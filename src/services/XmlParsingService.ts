/**
 * XML Parsing Service for EWS (Exchange Web Services) responses
 * Based on Microsoft Exchange Web Services XML Schema
 * https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/ews-xml-elements-in-exchange
 */

export interface EwsNamespaces {
    readonly SOAP_ENVELOPE: 'http://schemas.xmlsoap.org/soap/envelope/';
    readonly EWS_MESSAGES: 'http://schemas.microsoft.com/exchange/services/2006/messages';
    readonly EWS_TYPES: 'http://schemas.microsoft.com/exchange/services/2006/types';
}

export const EWS_NAMESPACES: EwsNamespaces = {
    SOAP_ENVELOPE: 'http://schemas.xmlsoap.org/soap/envelope/',
    EWS_MESSAGES: 'http://schemas.microsoft.com/exchange/services/2006/messages',
    EWS_TYPES: 'http://schemas.microsoft.com/exchange/services/2006/types'
} as const;

export interface ParsedEmail {
    id: string;
    subject: string;
    dateTimeSent: string;
    conversationId: string;
    body: { content: string };
    from: { emailAddress: { address: string } };
    toRecipients: Array<{ emailAddress: { address: string } }>;
    ccRecipients?: Array<{ emailAddress: { address: string } }>;
}

export interface ParsedThreadMessage {
    id: string;
    changeKey?: string;
    subject: string;
    from: string;
    to: string[];
    sentDate: Date;
    body: string;
    isFromCurrentUser: boolean;
}

export class XmlParsingService {
    
    /**
     * Parse EWS FindItem response to extract email messages
     * Schema: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/finditemresponse
     */
    public parseFindItemResponse(xmlResponse: string): ParsedEmail[] {
        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlResponse, 'text/xml');
            
            // Check for XML parsing errors
            const parserError = xmlDoc.querySelector('parsererror');
            if (parserError) {
                console.error('XML Parser Error:', parserError.textContent);
                return [];
            }

            // Use proper namespace-aware element selection
            const messages = this.getElementsByTagNameNS(
                xmlDoc, 
                EWS_NAMESPACES.EWS_TYPES, 
                'Message'
            );

            const emails: ParsedEmail[] = [];
            
            for (let i = 0; i < messages.length; i++) {
                const message = messages[i];
                const email = this.parseEmailMessage(message);
                if (email) {
                    emails.push(email);
                }
            }

            return emails;
        } catch (error) {
            console.error('Error parsing FindItem response:', error);
            return [];
        }
    }

    /**
     * Parse EWS GetConversationItems response to extract thread messages
     * Schema: https://docs.microsoft.com/en-us/exchange/client-developer/web-service-reference/getconversationitemsresponse
     */
    public parseConversationResponse(xmlResponse: string, currentUserEmail: string): ParsedThreadMessage[] {
        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlResponse, 'text/xml');
            
            const parserError = xmlDoc.querySelector('parsererror');
            if (parserError) {
                console.error('XML Parser Error:', parserError.textContent);
                return [];
            }

            const messages = this.getElementsByTagNameNS(
                xmlDoc,
                EWS_NAMESPACES.EWS_TYPES,
                'Message'
            );

            const threadMessages: ParsedThreadMessage[] = [];

            for (let i = 0; i < messages.length; i++) {
                const message = messages[i];
                const threadMessage = this.parseThreadMessage(message, currentUserEmail);
                if (threadMessage) {
                    threadMessages.push(threadMessage);
                }
            }

            // Sort by sent date
            return threadMessages.sort((a, b) => a.sentDate.getTime() - b.sentDate.getTime());
        } catch (error) {
            console.error('Error parsing conversation response:', error);
            return [];
        }
    }

    /**
     * Parse individual email message from EWS XML
     */
    private parseEmailMessage(messageElement: Element): ParsedEmail | null {
        try {
            const email: ParsedEmail = {
                id: this.getElementAttribute(messageElement, 'ItemId', 'Id') || '',
                subject: this.getElementTextContent(messageElement, 'Subject') || '',
                dateTimeSent: this.getElementTextContent(messageElement, 'DateTimeSent') || '',
                conversationId: this.getElementAttribute(messageElement, 'ConversationId', 'Id') || '',
                body: { 
                    content: this.cleanBodyContent(
                        this.getElementTextContent(messageElement, 'Body') || ''
                    ) 
                },
                from: { 
                    emailAddress: { 
                        address: this.getNestedElementText(messageElement, ['From', 'Mailbox', 'EmailAddress']) || ''
                    } 
                },
                toRecipients: this.parseRecipientsList(messageElement, 'ToRecipients'),
                ccRecipients: this.parseRecipientsList(messageElement, 'CcRecipients')
            };

            return email;
        } catch (error) {
            console.error('Error parsing email message:', error);
            return null;
        }
    }

    /**
     * Parse individual thread message from conversation XML
     */
    private parseThreadMessage(messageElement: Element, currentUserEmail: string): ParsedThreadMessage | null {
        try {
            const fromEmail = this.getNestedElementText(messageElement, ['From', 'Mailbox', 'EmailAddress']) || '';
            const dateTimeText = this.getElementTextContent(messageElement, 'DateTimeSent') || '';
            
            const threadMessage: ParsedThreadMessage = {
                id: this.getElementAttribute(messageElement, 'ItemId', 'Id') || '',
                changeKey: this.getElementAttribute(messageElement, 'ItemId', 'ChangeKey') || undefined,
                subject: this.getElementTextContent(messageElement, 'Subject') || '',
                from: fromEmail,
                to: this.parseRecipientsList(messageElement, 'ToRecipients').map(r => r.emailAddress.address),
                sentDate: dateTimeText ? new Date(dateTimeText) : new Date(),
                body: this.cleanBodyContent(this.getElementTextContent(messageElement, 'Body') || ''),
                isFromCurrentUser: fromEmail === currentUserEmail
            };

            return threadMessage;
        } catch (error) {
            console.error('Error parsing thread message:', error);
            return null;
        }
    }

    /**
     * Parse recipients list (ToRecipients, CcRecipients, BccRecipients)
     */
    private parseRecipientsList(messageElement: Element, recipientsType: string): Array<{ emailAddress: { address: string } }> {
        try {
            const recipientsElement = this.getChildElementByTagName(messageElement, recipientsType);
            if (!recipientsElement) {
                return [];
            }

            const mailboxElements = this.getChildElementsByTagName(recipientsElement, 'Mailbox');
            const recipients: Array<{ emailAddress: { address: string } }> = [];

            for (let i = 0; i < mailboxElements.length; i++) {
                const mailbox = mailboxElements[i];
                const emailAddress = this.getElementTextContent(mailbox, 'EmailAddress');
                if (emailAddress) {
                    recipients.push({ emailAddress: { address: emailAddress } });
                }
            }

            return recipients;
        } catch (error) {
            console.error(`Error parsing ${recipientsType}:`, error);
            return [];
        }
    }

    /**
     * Get element text content with namespace support
     */
    private getElementTextContent(parent: Element, localName: string): string {
        const element = this.getChildElementByTagName(parent, localName);
        return element?.textContent?.trim() || '';
    }

    /**
     * Get element attribute with namespace support
     */
    private getElementAttribute(parent: Element, localName: string, attributeName: string): string {
        const element = this.getChildElementByTagName(parent, localName);
        return element?.getAttribute(attributeName) || '';
    }

    /**
     * Get nested element text (e.g., From -> Mailbox -> EmailAddress)
     */
    private getNestedElementText(parent: Element, path: string[]): string {
        let currentElement: Element | null = parent;
        
        for (const localName of path) {
            if (!currentElement) break;
            currentElement = this.getChildElementByTagName(currentElement, localName);
        }
        
        return currentElement?.textContent?.trim() || '';
    }

    /**
     * Get child element by tag name (namespace-aware)
     */
    private getChildElementByTagName(parent: Element, localName: string): Element | null {
        // Try with EWS types namespace first
        const namespacedElements = this.getElementsByTagNameNS(parent, EWS_NAMESPACES.EWS_TYPES, localName);
        if (namespacedElements.length > 0) {
            return namespacedElements[0];
        }

        // Fallback to getElementsByTagName for local name
        const elements = parent.getElementsByTagName(localName);
        if (elements.length > 0) {
            return elements[0];
        }

        // Try with 't:' prefix as last resort
        const prefixedElements = parent.getElementsByTagName(`t:${localName}`);
        return prefixedElements.length > 0 ? prefixedElements[0] : null;
    }

    /**
     * Get child elements by tag name (namespace-aware)
     */
    private getChildElementsByTagName(parent: Element, localName: string): Element[] {
        // Try with EWS types namespace first
        const namespacedElements = this.getElementsByTagNameNS(parent, EWS_NAMESPACES.EWS_TYPES, localName);
        if (namespacedElements.length > 0) {
            return Array.from(namespacedElements);
        }

        // Fallback to getElementsByTagName for local name
        const elements = parent.getElementsByTagName(localName);
        if (elements.length > 0) {
            return Array.from(elements);
        }

        // Try with 't:' prefix as last resort
        const prefixedElements = parent.getElementsByTagName(`t:${localName}`);
        return Array.from(prefixedElements);
    }

    /**
     * Namespace-aware element selection
     */
    private getElementsByTagNameNS(parent: Document | Element, namespace: string, localName: string): Element[] {
        if (parent.getElementsByTagNameNS) {
            return Array.from(parent.getElementsByTagNameNS(namespace, localName));
        }

        // Fallback for browsers that don't support getElementsByTagNameNS
        const allElements = parent.getElementsByTagName('*');
        const matchingElements: Element[] = [];

        for (let i = 0; i < allElements.length; i++) {
            const element = allElements[i];
            if (element.localName === localName && element.namespaceURI === namespace) {
                matchingElements.push(element);
            }
        }

        return matchingElements;
    }

    /**
     * Clean HTML content from email body
     */
    private cleanBodyContent(content: string): string {
        if (!content) return '';

        // Remove CDATA wrapper if present
        const cdataMatch = content.match(/^<!\[CDATA\[([\s\S]*)\]\]>$/);
        if (cdataMatch) {
            content = cdataMatch[1];
        }

        // Strip HTML tags for plain text content
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = content;
        return tempDiv.textContent || tempDiv.innerText || '';
    }

    /**
     * Validate EWS response for errors
     */
    public validateEwsResponse(xmlResponse: string): { isValid: boolean; error?: string } {
        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlResponse, 'text/xml');
            
            const parserError = xmlDoc.querySelector('parsererror');
            if (parserError) {
                return { 
                    isValid: false, 
                    error: `XML Parse Error: ${parserError.textContent}` 
                };
            }

            // Check for SOAP faults
            const soapFault = xmlDoc.getElementsByTagNameNS(EWS_NAMESPACES.SOAP_ENVELOPE, 'Fault')[0];
            if (soapFault) {
                const faultString = soapFault.querySelector('faultstring')?.textContent || 'Unknown SOAP fault';
                return { 
                    isValid: false, 
                    error: `SOAP Fault: ${faultString}` 
                };
            }

            // Check for EWS response errors
            const responseMessages = xmlDoc.getElementsByTagNameNS(EWS_NAMESPACES.EWS_MESSAGES, 'ResponseMessages')[0];
            if (responseMessages) {
                const errorResponse = responseMessages.querySelector('[ResponseClass="Error"]');
                if (errorResponse) {
                    const responseCode = errorResponse.querySelector('ResponseCode')?.textContent || 'Unknown error';
                    const messageText = errorResponse.querySelector('MessageText')?.textContent || '';
                    return { 
                        isValid: false, 
                        error: `EWS Error ${responseCode}: ${messageText}` 
                    };
                }
            }

            return { isValid: true };
        } catch (error) {
            return { 
                isValid: false, 
                error: `Validation error: ${(error as Error).message}` 
            };
        }
    }
}
