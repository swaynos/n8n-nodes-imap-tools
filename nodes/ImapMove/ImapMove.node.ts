import type {
    ICredentialDataDecryptedObject,
    IDataObject,
    IExecuteFunctions,
    INodeExecutionData,
    INodeType,
    INodeTypeDescription,
} from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
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
            let rawMessageId = (this.getNodeParameter('messageId', i) as string).trim();
            const sourceMailbox = this.getNodeParameter('sourceMailbox', i) as string;
            const mailbox = this.getNodeParameter('mailbox', i) as string;

            let normalizedMessageId: string | undefined;
            let client: ImapFlow | undefined;
            let connected = false;
            let shouldContinue = false;
            let errorOutput: IDataObject | null = null;
            let successOutput: IDataObject | null = null;

            try {
                const credentials = (await this.getCredentials('imap')) as
                    | ICredentialDataDecryptedObject
                    | null;

                if (!credentials) {
                    throw new NodeOperationError(this.getNode(), 'IMAP credentials are required but missing.', {
                        itemIndex: i,
                    });
                }

                const host = (credentials.host as string) || '';
                const port = credentials.port !== undefined ? Number(credentials.port) : 993;
                const secure = credentials.secure !== false;
                const user = (credentials.user as string) || '';
                const password = (credentials.password as string) || '';
                const allowUnauthorizedCerts = credentials.allowUnauthorizedCerts === true;

                if (!rawMessageId) {
                    throw new NodeOperationError(this.getNode(), 'Message ID must be provided.', {
                        itemIndex: i,
                    });
                }

                if (!sourceMailbox) {
                    throw new NodeOperationError(this.getNode(), 'Source mailbox is required.', {
                        itemIndex: i,
                    });
                }

                if (!mailbox) {
                    throw new NodeOperationError(this.getNode(), 'Target mailbox is required.', {
                        itemIndex: i,
                    });
                }

                if (!Number.isFinite(port) || port <= 0) {
                    throw new NodeOperationError(
                        this.getNode(),
                        'IMAP credential port must be a positive number.',
                        { itemIndex: i },
                    );
                }

                if (!host || !user || !password) {
                    throw new NodeOperationError(
                        this.getNode(),
                        'IMAP credentials must include host, user, and password.',
                        { itemIndex: i },
                    );
                }

                normalizedMessageId = rawMessageId.startsWith('<') && rawMessageId.endsWith('>')
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

                client = new ImapFlow(clientOptions);

                await client.connect();
                connected = true;
                await client.mailboxOpen(sourceMailbox);

                const searchQuery: SearchObject = {
                    header: { 'message-id': normalizedMessageId },
                };

                const searchResult = await client.search(searchQuery, { uid: true });

                if (!searchResult || searchResult.length === 0) {
                    throw new NodeOperationError(
                        this.getNode(),
                        `Message with Message-ID ${normalizedMessageId} was not found in mailbox "${sourceMailbox}".`,
                        { itemIndex: i },
                    );
                }

                const messageUid = searchResult[0];

                const moveResult = await client.messageMove(String(messageUid), mailbox, {
                    uid: true,
                });
                if (moveResult === false) {
                    throw new NodeOperationError(
                        this.getNode(),
                        `Message with Message-ID ${normalizedMessageId} could not be moved to mailbox "${mailbox}".`,
                        { itemIndex: i },
                    );
                }

                successOutput = {
                    messageId: rawMessageId,
                    normalizedMessageId,
                    moved: true,
                    movedCount: 1,
                    matchedCount: searchResult.length,
                    movedUid: messageUid,
                    sourceMailbox,
                    targetMailbox: mailbox,
                };
            } catch (error) {
                const errorObject = error instanceof Error ? error : new Error('An unknown error occurred.');

                if (this.continueOnFail()) {
                    shouldContinue = true;
                    errorOutput = {
                        messageId: rawMessageId,
                        normalizedMessageId,
                        moved: false,
                        sourceMailbox,
                        targetMailbox: mailbox,
                        error: errorObject.message,
                    };
                } else {
                    if (error instanceof NodeOperationError) {
                        throw error;
                    }
                    throw new NodeOperationError(this.getNode(), errorObject, { itemIndex: i });
                }
            } finally {
                if (client) {
                    try {
                        if (connected) {
                            await client.logout();
                        } else {
                            await client.close();
                        }
                    } catch (closeError) {
                        const closeErrorMessage =
                            closeError instanceof Error
                                ? closeError.message
                                : 'Failed to close IMAP connection.';

                        if (this.continueOnFail()) {
                            shouldContinue = true;
                            if (errorOutput) {
                                const existingError = errorOutput.error as string | undefined;
                                errorOutput.error = existingError
                                    ? `${existingError}; ${closeErrorMessage}`
                                    : closeErrorMessage;
                            } else {
                                errorOutput = {
                                    messageId: rawMessageId,
                                    normalizedMessageId,
                                    moved: false,
                                    sourceMailbox,
                                    targetMailbox: mailbox,
                                    error: closeErrorMessage,
                                };
                            }
                        } else {
                            throw new NodeOperationError(this.getNode(), closeErrorMessage, {
                                itemIndex: i,
                            });
                        }
                    }
                }
            }

            if (shouldContinue) {
                returnData.push({
                    json: errorOutput ?? {
                        messageId: rawMessageId,
                        normalizedMessageId,
                        moved: false,
                        sourceMailbox,
                        targetMailbox: mailbox,
                    },
                    pairedItem: { item: i },
                });
                continue;
            }

            if (successOutput) {
                returnData.push({
                    json: successOutput,
                    pairedItem: { item: i },
                });
            }
        }

        return [returnData];
    }
}
