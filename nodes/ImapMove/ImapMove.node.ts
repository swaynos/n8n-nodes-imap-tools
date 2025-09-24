import { IExecuteFunctions } from 'n8n-core';
import { INodeExecutionData, INodeType, INodeTypeDescription } from 'n8n-workflow';
import { ImapFlow } from 'imapflow';

export class ImapMove implements INodeType {
    description: INodeTypeDescription = {
        displayName: 'IMAP Move',
        name: 'imapMove',
        group: ['transform'],
        version: 1,
        description: 'Move an email to a target mailbox',
        defaults: { name: 'IMAP Move' },
        inputs: ['main'],
        outputs: ['main'],
        properties: [
            {
                displayName: 'IMAP Host',
                name: 'host',
                type: 'string',
                default: 'imap.gmail.com',
            },
            {
                displayName: 'Port',
                name: 'port',
                type: 'number',
                default: 993,
            },
            {
                displayName: 'Secure',
                name: 'secure',
                type: 'boolean',
                default: true,
            },
            {
                displayName: 'User',
                name: 'user',
                type: 'string',
                default: '',
            },
            {
                displayName: 'Password / App Password',
                name: 'password',
                type: 'string',
                typeOptions: { password: true },
                default: '',
            },
            {
                displayName: 'Message UID',
                name: 'uid',
                type: 'number',
                default: 0,
                description: 'UID of the message to move',
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
            const host = this.getNodeParameter('host', i) as string;
            const port = this.getNodeParameter('port', i) as number;
            const secure = this.getNodeParameter('secure', i) as boolean;
            const user = this.getNodeParameter('user', i) as string;
            const password = this.getNodeParameter('password', i) as string;
            const uid = this.getNodeParameter('uid', i) as number;
            const mailbox = this.getNodeParameter('mailbox', i) as string;

            const client = new ImapFlow({
                host,
                port,
                secure,
                auth: { user, pass: password },
            });

            await client.connect();
            await client.mailboxOpen('INBOX'); // or wherever you expect the UID to reside
            await client.messageMove(uid, mailbox);
            await client.logout();

            returnData.push({ json: { uid, moved: true, mailbox } });
        }
        return [returnData];
    }
}