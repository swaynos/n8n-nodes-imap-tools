import type {
    ICredentialDataDecryptedObject,
    IDataObject,
    IExecuteFunctions,
    INodeExecutionData,
    INodeType,
    INodeTypeDescription,
} from 'n8n-workflow';
import { ImapFlow } from 'imapflow';
import type { ImapFlowOptions, SearchObject } from 'imapflow';

export class ImapMove implements INodeType {
    description: INodeTypeDescription = {
        displayName: 'IMAP Move',
        name: 'ImapMove',
        group: ['transform'],
        version: 1,
        description: 'Move an email to a target mailbox',
        defaults: { name: 'IMAP Move' },
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
                description: 'Value of the Message-ID header used to locate the email',
            },
            {
                displayName: 'Source Mailbox',
                name: 'sourceMailbox',
                type: 'string',
                default: 'INBOX',
                description: 'Mailbox where the message currently resides',
            },
            {
                displayName: 'Target Mailbox',
                name: 'mailbox',
                type: 'string',
                default: '[Gmail]/Spam',
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
            const sourceMailbox = this.getNodeParameter('sourceMailbox', i) as string;
            const mailbox = this.getNodeParameter('mailbox', i) as string;

            if (!rawMessageId) {
                throw new Error('Message ID must be provided.');
            }

            if (!sourceMailbox) {
                throw new Error('Source mailbox is required.');
            }

            if (!mailbox) {
                throw new Error('Target mailbox is required.');
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
                await client.mailboxOpen(sourceMailbox);

                const searchQuery: SearchObject = {
                    header: { 'message-id': normalizedMessageId },
                };

                const searchResult = await client.search(searchQuery, { uid: true });

                if (!searchResult || searchResult.length === 0) {
                    throw new Error(
                        `Message with Message-ID ${normalizedMessageId} was not found in mailbox "${sourceMailbox}".`,
                    );
                }

                const messageUid = searchResult[0];

                const moveResult = await client.messageMove(String(messageUid), mailbox, {
                    uid: true,
                });
                if (moveResult === false) {
                    throw new Error(
                        `Message with Message-ID ${normalizedMessageId} could not be moved to mailbox "${mailbox}".`,
                    );
                }

                returnData.push({
                    json: {
                        messageId: rawMessageId,
                        normalizedMessageId,
                        moved: true,
                        movedCount: 1,
                        matchedCount: searchResult.length,
                        movedUid: messageUid,
                        sourceMailbox,
                        targetMailbox: mailbox,
                    } as IDataObject,
                });
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
