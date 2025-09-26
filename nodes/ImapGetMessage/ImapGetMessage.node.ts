import type {
    ICredentialDataDecryptedObject,
    IDataObject,
    IExecuteFunctions,
    INodeExecutionData,
    INodeType,
    INodeTypeDescription,
} from 'n8n-workflow';
import { ImapFlow } from 'imapflow';
import type { ImapFlowOptions, FetchQueryObject, SearchObject } from 'imapflow';

export class ImapGetMessage implements INodeType {
    description: INodeTypeDescription = {
        displayName: 'IMAP Get Message',
        name: 'ImapGetMessage',
        group: ['transform'],
        version: 1,
        description: 'Retrieve a single email by Message-ID from a mailbox',
        defaults: { name: 'IMAP Get Message' },
        inputs: ['main'],
        outputs: ['main'],
        credentials: [
            {
                name: 'imap',
                required: true,
            },
        ],
        properties: [
            {
                displayName: 'Message ID',
                name: 'messageId',
                type: 'string',
                default: '',
                placeholder: '<message-id@example.com>',
                description: 'Value of the Message-ID header for the email to retrieve',
            },
            {
                displayName: 'Mailbox',
                name: 'mailbox',
                type: 'string',
                default: 'INBOX',
                description: 'Mailbox where the message resides',
            },
            {
                displayName: 'Include Raw Message',
                name: 'includeRaw',
                type: 'boolean',
                default: true,
                description:
                    'Whether to include the full raw message (base64 encoded) in the response',
            },
        ],
    };

    async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
        const items = this.getInputData();
        const returnData: INodeExecutionData[] = [];

        for (let i = 0; i < items.length; i++) {
            const credentials = (await this.getCredentials('imap')) as
                | ICredentialDataDecryptedObject
                | null;

            if (!credentials) {
                throw new Error('IMAP credentials are required but missing.');
            }

            const host = (credentials.host as string) || '';
            const port = credentials.port !== undefined ? Number(credentials.port) : 993;
            const secure = credentials.secure !== false;
            const user = (credentials.user as string) || '';
            const password = (credentials.password as string) || '';
            const allowUnauthorizedCerts = credentials.allowUnauthorizedCerts === true;
            const rawMessageId = (this.getNodeParameter('messageId', i) as string).trim();
            const mailbox = this.getNodeParameter('mailbox', i) as string;
            const includeRaw = this.getNodeParameter('includeRaw', i) as boolean;

            if (!rawMessageId) {
                throw new Error('Message ID must be provided.');
            }

            if (!mailbox) {
                throw new Error('Mailbox is required to fetch a message.');
            }

            if (!Number.isFinite(port) || port <= 0) {
                throw new Error('IMAP credential port must be a positive number.');
            }

            if (!host || !user || !password) {
                throw new Error('IMAP credentials must include host, user, and password.');
            }

            const normalizedMessageId = rawMessageId.startsWith('<') && rawMessageId.endsWith('>')
                ? rawMessageId
                : `<${rawMessageId.replace(/^<|>$/g, '')}>`;

            const clientOptions: ImapFlowOptions = {
                host,
                port,
                secure,
                auth: { user, pass: password },
            };

            if (allowUnauthorizedCerts) {
                clientOptions.tls = { rejectUnauthorized: false };
            }

            const client = new ImapFlow(clientOptions);

            let connected = false;
            try {
                await client.connect();
                connected = true;

                await client.mailboxOpen(mailbox, { readOnly: true });

                const searchQuery: SearchObject = {
                    header: { 'message-id': normalizedMessageId },
                };

                const searchResult = await client.search(searchQuery, { uid: true });
                if (!searchResult || searchResult.length === 0) {
                    throw new Error(
                        `Message with Message-ID ${normalizedMessageId} was not found in mailbox "${mailbox}".`,
                    );
                }

                const fetchQuery: FetchQueryObject = {
                    uid: true,
                    flags: true,
                    envelope: true,
                    internalDate: true,
                    size: true,
                    threadId: true,
                    labels: true,
                    headers: true,
                    bodyStructure: true,
                    source: includeRaw,
                };

                const message = await client.fetchOne(String(searchResult[0]), fetchQuery, { uid: true });

                if (!message) {
                    throw new Error(
                        `Message with Message-ID ${normalizedMessageId} was not found in mailbox "${mailbox}".`,
                    );
                }

                const output: IDataObject = {
                    messageId: rawMessageId,
                    normalizedMessageId,
                    mailbox,
                    uid: message.uid,
                    seq: message.seq,
                    flags: Array.from(message.flags ?? []),
                };

                if (message.size !== undefined) {
                    output.size = message.size;
                }

                if (message.labels && message.labels.size > 0) {
                    output.labels = Array.from(message.labels);
                }

                if (message.internalDate) {
                    output.internalDate =
                        message.internalDate instanceof Date
                            ? message.internalDate.toISOString()
                            : message.internalDate;
                }

                if (message.envelope) {
                    const { date, ...rest } = message.envelope;
                    if (date) {
                        output.envelope = {
                            ...rest,
                            date: date instanceof Date ? date.toISOString() : date,
                        } as IDataObject;
                    } else {
                        output.envelope = { ...rest } as IDataObject;
                    }
                }

                if (message.bodyStructure) {
                    output.bodyStructure = message.bodyStructure as unknown as IDataObject;
                }

                if (message.headers) {
                    output.headers = message.headers.toString('utf8');
                }

                if (typeof message.modseq === 'bigint') {
                    output.modseq = message.modseq.toString();
                }

                if (message.emailId) {
                    output.emailId = message.emailId;
                }

                if (message.threadId) {
                    output.threadId = message.threadId;
                }

                if (message.id) {
                    output.id = message.id;
                }

                if (includeRaw && message.source) {
                    output.source = message.source.toString('base64');
                    output.sourceEncoding = 'base64';
                    output.sourceBytes = message.source.length;
                }

                returnData.push({ json: output });
            } finally {
                if (connected) {
                    await client.logout();
                } else {
                    await client.close();
                }
            }
        }

        return [returnData];
    }
}
