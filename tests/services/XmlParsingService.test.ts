import { XmlParsingService } from "../../src/services/XmlParsingService";

describe("XmlParsingService", () => {
  let xmlParsingService: XmlParsingService;

  beforeEach(() => {
    xmlParsingService = new XmlParsingService();
  });

  describe("EWS FindItem Response Parsing", () => {
    const mockFindItemResponse = `<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
    <s:Body>
        <m:FindItemResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <m:ResponseMessages>
                <m:FindItemResponseMessage ResponseClass="Success">
                    <m:ResponseCode>NoError</m:ResponseCode>
                    <m:RootFolder TotalItemsInView="1" IncludesLastItemInRange="true">
                        <t:Items>
                            <t:Message>
                                <t:ItemId Id="test-id" ChangeKey="test-key"/>
                                <t:Subject>Test Subject</t:Subject>
                                <t:DateTimeSent>2024-01-15T10:30:00Z</t:DateTimeSent>
                                <t:ConversationId Id="conv-id"/>
                                <t:Body BodyType="Text">Test email body</t:Body>
                                <t:From>
                                    <t:Mailbox>
                                        <t:EmailAddress>sender@example.com</t:EmailAddress>
                                    </t:Mailbox>
                                </t:From>
                                <t:ToRecipients>
                                    <t:Mailbox>
                                        <t:EmailAddress>recipient@example.com</t:EmailAddress>
                                    </t:Mailbox>
                                </t:ToRecipients>
                            </t:Message>
                        </t:Items>
                    </m:RootFolder>
                </m:FindItemResponseMessage>
            </m:ResponseMessages>
        </m:FindItemResponse>
    </s:Body>
</s:Envelope>`;

    it("should parse FindItem response and extract emails correctly", () => {
      const emails =
        xmlParsingService.parseFindItemResponse(mockFindItemResponse);

      expect(emails).toHaveLength(1);

      const email = emails[0];
      expect(email.subject).toBe("Test Subject");
      expect(email.dateTimeSent).toBe("2024-01-15T10:30:00Z");
      expect(email.from.emailAddress.address).toBe("sender@example.com");
      expect(email.toRecipients).toHaveLength(1);
      expect(email.toRecipients[0].emailAddress.address).toBe(
        "recipient@example.com",
      );
      expect(email.body.content).toBe("Test email body");
    });

    it("should handle empty FindItem response", () => {
      const emptyResponse = `<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
    <s:Body>
        <m:FindItemResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <m:ResponseMessages>
                <m:FindItemResponseMessage ResponseClass="Success">
                    <m:ResponseCode>NoError</m:ResponseCode>
                    <m:RootFolder TotalItemsInView="0" IncludesLastItemInRange="true">
                        <t:Items/>
                    </m:RootFolder>
                </m:FindItemResponseMessage>
            </m:ResponseMessages>
        </m:FindItemResponse>
    </s:Body>
</s:Envelope>`;

      const emails = xmlParsingService.parseFindItemResponse(emptyResponse);
      expect(emails).toHaveLength(0);
    });
  });

  describe("EWS Response Validation", () => {
    it("should validate valid EWS response", () => {
      const validResponse = `<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
    <s:Body>
        <m:FindItemResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
            <m:ResponseMessages>
                <m:FindItemResponseMessage ResponseClass="Success">
                    <m:ResponseCode>NoError</m:ResponseCode>
                </m:FindItemResponseMessage>
            </m:ResponseMessages>
        </m:FindItemResponse>
    </s:Body>
</s:Envelope>`;

      const validation = xmlParsingService.validateEwsResponse(validResponse);
      expect(validation.isValid).toBe(true);
      expect(validation.error).toBeUndefined();
    });

    it("should detect XML parse errors", () => {
      const invalidXml = `<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
    <s:Body>
        <unclosed-tag>
    </s:Body>
</s:Envelope>`;

      const validation = xmlParsingService.validateEwsResponse(invalidXml);
      expect(validation.isValid).toBe(false);
      expect(validation.error).toContain("XML Parse Error");
    });
  });

  describe("GetConversationItems Response Parsing", () => {
    const conversationResponse = `<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
    <s:Body>
        <m:GetConversationItemsResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            <m:ResponseMessages>
                <m:GetConversationItemsResponseMessage ResponseClass="Success">
                    <m:ResponseCode>NoError</m:ResponseCode>
                    <m:Conversation>
                        <t:ConversationId Id="test-conv"/>
                        <t:ConversationNodes>
                            <t:ConversationNode>
                                <t:Items>
                                    <t:Message>
                                        <t:ItemId Id="msg1" ChangeKey="ck1"/>
                                        <t:Subject>Original Message</t:Subject>
                                        <t:DateTimeSent>2024-01-15T09:00:00Z</t:DateTimeSent>
                                        <t:Body BodyType="Text">First message</t:Body>
                                        <t:From>
                                            <t:Mailbox>
                                                <t:EmailAddress>user@company.com</t:EmailAddress>
                                            </t:Mailbox>
                                        </t:From>
                                        <t:ToRecipients>
                                            <t:Mailbox>
                                                <t:EmailAddress>recipient@example.com</t:EmailAddress>
                                            </t:Mailbox>
                                        </t:ToRecipients>
                                    </t:Message>
                                </t:Items>
                            </t:ConversationNode>
                        </t:ConversationNodes>
                    </m:Conversation>
                </m:GetConversationItemsResponseMessage>
            </m:ResponseMessages>
        </m:GetConversationItemsResponse>
    </s:Body>
</s:Envelope>`;

    it("should parse conversation response correctly", () => {
      const currentUserEmail = "user@company.com";
      const threadMessages = xmlParsingService.parseGetConversationItemsResponse(
        conversationResponse,
        currentUserEmail,
      );

      expect(threadMessages).toHaveLength(1);

      const message = threadMessages[0];
      expect(message.subject).toBe("Original Message");
      expect(message.from).toBe("user@company.com");
      expect(message.isFromCurrentUser).toBe(true);
      expect(message.body).toBe("First message");
      expect(message.to).toContain("recipient@example.com");
    });
  });
});
