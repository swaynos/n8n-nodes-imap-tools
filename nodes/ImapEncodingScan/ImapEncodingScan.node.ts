import type {
    INodeType,
    INodeTypeDescription,
    IExecuteFunctions,
    IDataObject,
    INodeExecutionData,
    ICredentialDataDecryptedObject,
} from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { ImapFlow } from 'imapflow';
import type {
    ImapFlowOptions,
    FetchQueryObject,
    MessageEnvelopeObject,
    MessageStructureObject,
    SearchObject,
} from 'imapflow';

type SearchCriteriaValue = string | number | boolean | null | SearchCriteriaValue[] | SearchObject;

function flattenSearchTokens(value: SearchCriteriaValue): string[] {
    const tokens: string[] = [];

    const recurse = (current: SearchCriteriaValue): void => {
        if (current === null || current === undefined) {
            return;
        }

        if (Array.isArray(current)) {
            for (const entry of current) {
                recurse(entry);
            }
            return;
        }

        if (typeof current === 'object') {
            throw new Error('Search criteria arrays may not contain nested objects.');
        }

        tokens.push(String(current));
    };

    recurse(value);

    return tokens;
}

function tokensToSearchObject(tokens: string[]): SearchObject {
    const search: SearchObject = {};

    const nextToken = (keyword: string): string => {
        const next = tokens.shift();
        if (next === undefined) {
            throw new Error(`${keyword} requires an additional value.`);
        }
        return next;
    };

    while (tokens.length) {
        const rawToken = tokens.shift();
        if (rawToken === undefined) {
            break;
        }
        const token = rawToken.toUpperCase();

        switch (token) {
            case 'ALL':
                search.all = true;
                break;
            case 'SEEN':
                search.seen = true;
                break;
            case 'UNSEEN':
                search.seen = false;
                break;
            case 'ANSWERED':
                search.answered = true;
                break;
            case 'UNANSWERED':
                search.answered = false;
                break;
            case 'DELETED':
                search.deleted = true;
                break;
            case 'UNDELETED':
                search.deleted = false;
                break;
            case 'DRAFT':
                search.draft = true;
                break;
            case 'UNDRAFT':
                search.draft = false;
                break;
            case 'FLAGGED':
                search.flagged = true;
                break;
            case 'UNFLAGGED':
                search.flagged = false;
                break;
            case 'RECENT':
                search.recent = true;
                break;
            case 'OLD':
                search.old = true;
                break;
            case 'NEW':
                search.new = true;
                break;
            case 'FROM':
                search.from = nextToken(token);
                break;
            case 'TO':
                search.to = nextToken(token);
                break;
            case 'CC':
                search.cc = nextToken(token);
                break;
            case 'BCC':
                search.bcc = nextToken(token);
                break;
            case 'SUBJECT':
                search.subject = nextToken(token);
                break;
            case 'BODY':
                search.body = nextToken(token);
                break;
            case 'LARGER':
                search.larger = Number(nextToken(token));
                break;
            case 'SMALLER':
                search.smaller = Number(nextToken(token));
                break;
            case 'BEFORE':
                search.before = nextToken(token);
                break;
            case 'ON':
                search.on = nextToken(token);
                break;
            case 'SINCE':
                search.since = nextToken(token);
                break;
            case 'SENTBEFORE':
                search.sentBefore = nextToken(token);
                break;
            case 'SENTON':
                search.sentOn = nextToken(token);
                break;
            case 'SENTSINCE':
                search.sentSince = nextToken(token);
                break;
            case 'UID':
                search.uid = nextToken(token);
                break;
            case 'HEADER': {
                const headerKey = nextToken(token);
                const headerValue = nextToken(token);
                const cleanKey = headerKey.toLowerCase();
                const parsedValue =
                    headerValue.toLowerCase() === 'true'
                        ? true
                        : headerValue.toLowerCase() === 'false'
                            ? false
                            : headerValue;
                search.header = { ...(search.header ?? {}), [cleanKey]: parsedValue };
                break;
            }
            default:
                throw new Error(`Unsupported search token "${rawToken}".`);
        }
    }

    return Object.keys(search).length > 0 ? search : { all: true };
}

function cleanEncoding(value: string | undefined | null): string {
    return value ? value.replace(/\s+/g, ' ').trim() : '';
}

function extractMessageId(
    envelope: MessageEnvelopeObject | undefined,
    rawHeaders: string | undefined,
): string | undefined {
    if (envelope?.messageId) {
        return envelope.messageId;
    }

    if (!rawHeaders) {
        return undefined;
    }

    const unfoldedHeaders = rawHeaders.replace(/\r?\n[ \t]+/g, ' ');
    const match = unfoldedHeaders.match(/^Message-ID:\s*(.+)$/im);
    return match?.[1]?.trim();
}

function describeStructure(node: MessageStructureObject): string {
    const parts: string[] = [];

    if (node.type) {
        parts.push(node.type);
    }

    if (node.part) {
        parts.push(`part ${node.part}`);
    }

    if (node.disposition) {
        parts.push(node.disposition);
    }

    return parts.length > 0 ? parts.join(' ') : 'message';
}

function collectEncodingMatches(
    structure: MessageStructureObject | undefined,
    regex: RegExp,
): string[] {
    if (!structure) {
        return [];
    }

    const matches: string[] = [];
    const stack: MessageStructureObject[] = [structure];

    while (stack.length > 0) {
        const node = stack.shift();
        if (!node) {
            continue;
        }

        const encoding = cleanEncoding(node.encoding);
        if (encoding && regex.test(encoding)) {
            matches.push(`${describeStructure(node)}: ${encoding}`);
        }

        if (node.childNodes && node.childNodes.length) {
            stack.push(...node.childNodes);
        }
    }

    return matches;
}

function collectHeaderEncodingMatches(rawHeaders: string | undefined, regex: RegExp): string[] {
    if (!rawHeaders) {
        return [];
    }

    const matches: string[] = [];
    const unfolded = rawHeaders.replace(/\r?\n[ \t]+/g, ' ');
    const headerRegex = /^Content-Transfer-Encoding:\s*(.+)$/gim;
    let match: RegExpExecArray | null;

    // Only flag content-transfer-encoding headers to avoid eager matches on unrelated fields.
    while ((match = headerRegex.exec(unfolded)) !== null) {
        const cleaned = cleanEncoding(match[1]);
        if (cleaned && regex.test(cleaned)) {
            matches.push(`headers: ${cleaned}`);
        }
    }

    return matches;
}

export class ImapEncodingScan implements INodeType {
    description: INodeTypeDescription = {
        displayName: 'IMAP Encoding Scan',
        name: 'ImapEncodingScan',
        group: ['transform'],
        version: 1,
        description:
            'Search for emails whose Content-Transfer-Encoding or raw headers match a regex pattern',
        defaults: { name: 'IMAP Encoding Scan' },
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
                displayName: 'Mailbox',
                name: 'mailbox',
                type: 'string',
                default: '[Gmail]/All Mail',
                description: 'Mailbox to scan for encoding issues',
            },
            {
                displayName: 'Search Criteria (JSON)',
                name: 'searchCriteria',
                type: 'string',
                typeOptions: { rows: 3 },
                default: '["UNSEEN"]',
                description:
                    'JSON describing the IMAP search. Provide either an array of IMAP tokens (e.g. ["UNSEEN","SINCE","01-Jan-2024"]) or an object that matches the imapflow SearchObject shape.',
            },
            {
                displayName: 'Encoding Regex Pattern',
                name: 'encodingPattern',
                type: 'string',
                default: '(?:AMAZONSES"?|8bit\\s*\\+)',
                description:
                    'Regular expression used against MIME Content-Transfer-Encoding values and raw headers to flag suspicious messages',
            },
            {
                displayName: 'Options',
                name: 'options',
                type: 'collection',
                default: {},
                placeholder: 'Add Option',
                options: [
                    {
                        displayName: 'Case Insensitive Pattern',
                        name: 'caseInsensitive',
                        type: 'boolean',
                        default: true,
                        description: 'Whether the regex should run in case-insensitive mode',
                    },
                    {
                        displayName: 'Stop After First Match',
                        name: 'stopAfterFirst',
                        type: 'boolean',
                        default: true,
                        description: 'Stop scanning immediately when the first match is found',
                    },
                    {
                        displayName: 'Max Results',
                        name: 'maxResults',
                        type: 'number',
                        typeOptions: { minValue: 0 },
                        default: 0,
                        description: 'Stop after collecting this many matches. Use 0 for no limit.',
                    },
                    {
                        displayName: 'Progress Interval',
                        name: 'progressEvery',
                        type: 'number',
                        typeOptions: { minValue: 0 },
                        default: 200,
                        description:
                            'Log progress after scanning this many messages. Set to 0 to disable logging.',
                    },
                    {
                        displayName: 'Include Raw Headers in Output',
                        name: 'includeRawHeaders',
                        type: 'boolean',
                        default: false,
                        description: 'Attach the raw header text to each match for downstream inspection',
                    },
                    {
                        displayName: 'Include Raw Message on Match',
                        name: 'includeRawMessage',
                        type: 'boolean',
                        default: false,
                        description:
                            'Fetch and attach the full raw message (base64 encoded) for matches only',
                    },
                    {
                        displayName: 'Raw Message Property Name',
                        name: 'rawMessageProperty',
                        type: 'string',
                        default: 'rawMessage',
                        description:
                            'Property name to use when attaching the raw message (only used when including raw messages)',
                    },
                ],
            },
        ],
    };

    async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
        const items = this.getInputData();
        const returnData: INodeExecutionData[] = [];

        for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
            const mailbox = (this.getNodeParameter('mailbox', itemIndex) as string).trim();
            const searchCriteriaRaw = this.getNodeParameter('searchCriteria', itemIndex) as string;
            const encodingPattern = (this.getNodeParameter('encodingPattern', itemIndex) as string).trim();
            const options = (this.getNodeParameter('options', itemIndex, {}) as IDataObject) ?? {};

            if (!mailbox) {
                throw new NodeOperationError(this.getNode(), 'Mailbox must be provided.', {
                    itemIndex,
                });
            }

            if (!searchCriteriaRaw) {
                throw new NodeOperationError(
                    this.getNode(),
                    'Search criteria must not be empty. Provide JSON describing the IMAP search.',
                    { itemIndex },
                );
            }

            if (!encodingPattern) {
                throw new NodeOperationError(this.getNode(), 'Encoding regex pattern is required.', {
                    itemIndex,
                });
            }

            let searchCriteriaInput: SearchCriteriaValue;
            try {
                searchCriteriaInput = JSON.parse(searchCriteriaRaw) as SearchCriteriaValue;
            } catch (error) {
                throw new NodeOperationError(
                    this.getNode(),
                    `Search criteria must be valid JSON. ${error instanceof Error ? error.message : ''}`.trim(),
                    { itemIndex },
                );
            }

            let searchQuery: SearchObject;
            if (Array.isArray(searchCriteriaInput)) {
                try {
                    const tokens = flattenSearchTokens(searchCriteriaInput);
                    searchQuery = tokensToSearchObject(tokens);
                } catch (error) {
                    throw new NodeOperationError(
                        this.getNode(),
                        error instanceof Error ? error.message : 'Failed to parse search criteria array.',
                        { itemIndex },
                    );
                }
            } else if (typeof searchCriteriaInput === 'object' && searchCriteriaInput !== null) {
                const searchObject = searchCriteriaInput as SearchObject;
                searchQuery = Object.keys(searchObject).length > 0 ? searchObject : { all: true };
            } else {
                throw new NodeOperationError(
                    this.getNode(),
                    'Search criteria JSON must describe either an array of IMAP tokens or an object.',
                    { itemIndex },
                );
            }

            const caseInsensitive = options.caseInsensitive !== false;
            const includeRawHeaders = options.includeRawHeaders === true;
            const stopAfterFirst = options.stopAfterFirst !== false;
            const includeRawMessage = options.includeRawMessage === true;
            const rawMessageProperty =
                typeof options.rawMessageProperty === 'string' && options.rawMessageProperty.trim()
                    ? options.rawMessageProperty.trim()
                    : 'rawMessage';
            const maxResults =
                typeof options.maxResults === 'number' && options.maxResults > 0
                    ? options.maxResults
                    : 0;
            const progressEvery =
                typeof options.progressEvery === 'number' && options.progressEvery > 0
                    ? options.progressEvery
                    : 0;

            let encodingRegex: RegExp;
            try {
                encodingRegex = new RegExp(encodingPattern, caseInsensitive ? 'i' : undefined);
            } catch (error) {
                throw new NodeOperationError(
                    this.getNode(),
                    `Invalid regex pattern: ${error instanceof Error ? error.message : String(error)}`,
                    { itemIndex },
                );
            }

            const credentials = (await this.getCredentials('imap')) as
                | ICredentialDataDecryptedObject
                | null;

            if (!credentials) {
                throw new NodeOperationError(this.getNode(), 'IMAP credentials are required but missing.', {
                    itemIndex,
                });
            }

            const host = (credentials.host as string) || '';
            const port =
                credentials.port !== undefined && Number.isFinite(Number(credentials.port))
                    ? Number(credentials.port)
                    : 993;
            const secure = credentials.secure !== false;
            const user = (credentials.user as string) || '';
            const password = (credentials.password as string) || '';
            const allowUnauthorizedCerts = credentials.allowUnauthorizedCerts === true;

            if (!host || !user || !password) {
                throw new NodeOperationError(
                    this.getNode(),
                    'IMAP credentials must include host, user, and password.',
                    { itemIndex },
                );
            }

            if (!Number.isFinite(port) || port <= 0) {
                throw new NodeOperationError(
                    this.getNode(),
                    'IMAP credential port must be a positive number.',
                    { itemIndex },
                );
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

                await client.mailboxOpen(mailbox, { readOnly: true });

                const searchResult = await client.search(searchQuery, { uid: true });
                const totalCandidates = Array.isArray(searchResult) ? searchResult.length : 0;

                if (!searchResult || searchResult.length === 0) {
                    const summary: IDataObject = {
                        matches: 0,
                        totalCandidates,
                        scanned: 0,
                        matchDetails: [],
                    };
                    returnData.push({
                        json: summary,
                        pairedItem: { item: itemIndex },
                    });
                    continue;
                }

                const fetchQuery: FetchQueryObject = {
                    uid: true,
                    envelope: true,
                    bodyStructure: true,
                    headers: true,
                };

                const matches: IDataObject[] = [];
                let scanned = 0;

                for (let idx = 0; idx < searchResult.length; idx++) {
                    const uid = searchResult[idx];
                    const message = await client.fetchOne(String(uid), fetchQuery, { uid: true });
                    scanned += 1;

                    if (!message || !message.uid) {
                        continue;
                    }

                    const suspectDetails = collectEncodingMatches(message.bodyStructure, encodingRegex);

                    let rawHeaders: string | undefined;
                    if (message.headers) {
                        rawHeaders = message.headers.toString('utf8');

                        if (suspectDetails.length === 0) {
                            suspectDetails.push(
                                ...collectHeaderEncodingMatches(rawHeaders, encodingRegex),
                            );
                        }
                    }

                    if (suspectDetails.length === 0) {
                        if (progressEvery && scanned % progressEvery === 0) {
                            const logMessage = `Scanned ${scanned} of ${totalCandidates} messages; ${matches.length} matches so far.`;
                            if (this.logger) {
                                this.logger.info(logMessage);
                            } else {
                                // eslint-disable-next-line no-console
                                console.log(logMessage);
                            }
                        }
                        continue;
                    }

                    const envelope = message.envelope;
                    const entry: IDataObject = {
                        imapUid: String(message.uid),
                        matchSources: Array.from(new Set(suspectDetails)),
                    };

                    if (envelope) {
                        if (envelope.subject) {
                            entry.subject = envelope.subject;
                        }
                        if (envelope.date) {
                            entry.date =
                                envelope.date instanceof Date
                                    ? envelope.date.toISOString()
                                    : envelope.date;
                        }
                        if (envelope.from && envelope.from.length > 0) {
                            entry.from = envelope.from.map((address) => address.address ?? address.name);
                        }
                    }

                    if (includeRawHeaders && rawHeaders) {
                        entry.rawHeaders = rawHeaders;
                    }

                    if (includeRawMessage) {
                        const rawMessage = await client.fetchOne(
                            String(uid),
                            { source: true, size: true },
                            { uid: true },
                        );
                        if (rawMessage && rawMessage.source) {
                            entry[rawMessageProperty] = rawMessage.source.toString('base64');
                            entry.rawMessageEncoding = 'base64';
                            entry.rawMessageBytes = rawMessage.source.length;
                        } else {
                            entry.rawMessageError = 'Raw message was not available for this UID.';
                        }
                    }

                    const messageId = extractMessageId(envelope, rawHeaders);
                    if (messageId) {
                        entry.messageId = messageId;
                    }

                    matches.push(entry);

                    if (progressEvery && scanned % progressEvery === 0) {
                        const logMessage = `Scanned ${scanned} of ${totalCandidates} messages; ${matches.length} matches so far.`;
                        if (this.logger) {
                            this.logger.info(logMessage);
                        } else {
                            // eslint-disable-next-line no-console
                            console.log(logMessage);
                        }
                    }

                    if (stopAfterFirst || (maxResults && matches.length >= maxResults)) {
                        break;
                    }
                }

                const summary: IDataObject = {
                    matches: matches.length,
                    totalCandidates,
                    scanned,
                    matchDetails: matches,
                };
                returnData.push({
                    json: summary,
                    pairedItem: { item: itemIndex },
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
