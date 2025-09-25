import type {
    ICredentialDataDecryptedObject,
    IExecuteFunctions,
    INodeExecutionData,
    INodeType,
    INodeTypeDescription,
} from 'n8n-workflow';
import { ImapFlow } from 'imapflow';
import type { ImapFlowOptions } from 'imapflow';

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
                displayName: 'Message UID',
                name: 'uid',
                type: 'number',
                default: 0,
                description: 'UID of the message to move',
            },
            {
                displayName: 'Source Mailbox',
                name: 'sourceMailbox',
                type: 'string',
                default: 'INBOX',
                description: 'Mailbox where the UID currently resides',
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
            const credentials = (await this.getCredentials(
                'imap',
            )) as ICredentialDataDecryptedObject | null;

            if (!credentials) {
                throw new Error('IMAP credentials are required but missing.');
            }
            const host = (credentials.host as string) || '';
            const port = credentials.port !== undefined ? Number(credentials.port) : 993;
            const secure = credentials.secure !== false;
            const user = (credentials.user as string) || '';
            const password = (credentials.password as string) || '';
            const allowUnauthorizedCerts = credentials.allowUnauthorizedCerts === true;
            const uid = this.getNodeParameter('uid', i) as number;
            const sourceMailbox = this.getNodeParameter('sourceMailbox', i) as string;
            const mailbox = this.getNodeParameter('mailbox', i) as string;

            if (!uid || uid <= 0) {
                throw new Error('Message UID must be a positive number.');
            }

            if (!Number.isFinite(port) || port <= 0) {
                throw new Error('IMAP credential port must be a positive number.');
            }

            if (!host || !user || !password) {
                throw new Error('IMAP credentials must include host, user, and password.');
            }

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
                const moveResult = await client.messageMove(String(uid), mailbox, { uid: true });
                if (moveResult === false) {
                    throw new Error(
                        `Message with UID ${uid} was not found in mailbox "${sourceMailbox}".`,
                    );
                }
            } finally {
                if (connected) {
                    await client.logout();
                } else {
                    await client.close();
                }
            }

            returnData.push({ json: { uid, moved: true, mailbox, sourceMailbox } });
        }
        return [returnData];
    }
}
