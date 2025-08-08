import { LlmService } from '../../src/services/LlmService';
import { RetryService } from '../../src/services/RetryService';
import { TaskpaneManager } from '../../src/taskpane/taskpane';
import { Configuration } from '../../src/models/Configuration';
import { FollowupEmail, ThreadMessage } from '../../src/models/FollowupEmail';

// Minimal DOM stubs for TaskpaneManager dependencies
function ensureDom() {
  if (!(global as any).document) {
    (global as any).document = {
      getElementById: (id: string) => {
        // Return minimal elements with needed props
        const el: any = {
          id,
          style: {},
          children: [],
          innerHTML: '',
          addEventListener: jest.fn(),
          appendChild: jest.fn(),
          querySelectorAll: () => [],
          querySelector: () => null,
          set textContent(v: string) { this._text = v; },
          get textContent() { return this._text; },
          value: '',
          disabled: false,
          checked: false,
          classList: { add: jest.fn(), remove: jest.fn(), toggle: jest.fn(), contains: () => false }
        };
        return el;
      },
      createElement: (tag: string) => ({
        tag,
        style: {},
        innerHTML: '',
        className: '',
        appendChild: jest.fn(),
        querySelectorAll: () => [],
        addEventListener: jest.fn()
      })
    } as any;
  }
}

// Helper to build configuration
const baseConfig: Configuration = {
  emailCount: 50,
  daysBack: 14,
  lastAnalysisDate: new Date(),
  enableLlmSummary: true,
  enableLlmSuggestions: true,
  llmApiEndpoint: 'https://api.test/openai',
  llmApiKey: 'key',
  llmModel: 'gpt-test',
  showSnoozedEmails: true,
  showDismissedEmails: true,
  snoozeOptions: [],
  selectedAccounts: []
};

describe('LlmService.healthCheck', () => {
  beforeEach(() => {
    (global.fetch as jest.Mock).mockReset();
  });

  it('returns true when OK appears in response', async () => {
    (global.fetch as jest.Mock).mockResolvedValue({
      ok: true,
      json: async () => ({ choices: [{ message: { content: 'OK' }, finish_reason: 'stop' }], usage: { prompt_tokens: 1, completion_tokens: 1, total_tokens: 2 }, model: 'x' })
    });
    const svc = new LlmService(baseConfig, new RetryService());
    const ok = await svc.healthCheck();
    expect(ok).toBe(true);
  });

  it('returns false on failure', async () => {
    (global.fetch as jest.Mock).mockResolvedValue({ ok: false, status: 500, statusText: 'err', text: async () => 'fail' });
    const svc = new LlmService(baseConfig, new RetryService());
    const ok = await svc.healthCheck();
    expect(ok).toBe(false);
  });
});

describe('Taskpane Reply/Forward logic', () => {
  beforeEach(() => {
    ensureDom();
    localStorage.clear();
  });

  function buildTaskpane(): TaskpaneManager {
    const tp = new TaskpaneManager();
    // Inject emails
    (tp as any).allEmails = [] as FollowupEmail[];
    return tp;
  }

  function setEmails(tp: TaskpaneManager, thread: ThreadMessage[]) {
    const email: FollowupEmail = {
      id: thread[thread.length - 1].id,
      subject: thread[thread.length - 1].subject,
      recipients: thread[thread.length - 1].to,
      sentDate: thread[thread.length - 1].sentDate,
      body: thread[thread.length - 1].body,
      summary: 's',
      priority: 'low',
      daysWithoutResponse: 1,
      conversationId: 'conv1',
      hasAttachments: false,
      accountEmail: thread[thread.length - 1].from,
      threadMessages: thread,
      isSnoozed: false,
      isDismissed: false
    };
    (tp as any).allEmails = [email];
    return email.id;
  }

  it('replyToEmail builds Reply All recipients excluding current user', async () => {
    const tp = buildTaskpane();
    const current = (global as any).Office.context.mailbox.userProfile.emailAddress.toLowerCase();
    const thread: ThreadMessage[] = [
      { id: 'm1', subject: 'Topic', from: 'alice@example.com', to: ['bob@example.com', current], sentDate: new Date(), body: 'Body', isFromCurrentUser: false }
    ];
    const id = setEmails(tp, thread);
  const mailbox = (global as any).Office.context.mailbox;
  mailbox.displayNewMessageForm = jest.fn();
  await (tp as any).replyToEmail(id);
  expect(mailbox.displayNewMessageForm).toHaveBeenCalled();
  const arg = (mailbox.displayNewMessageForm as jest.Mock).mock.calls[0][0] as any;
  expect(arg.subject.startsWith('Re:')).toBe(true);
  expect(arg.toRecipients).toContain('alice@example.com');
  expect(arg.toRecipients).not.toContain(current);
  });

  it('forwardEmail prefixes FW and leaves empty toRecipients', async () => {
    const tp = buildTaskpane();
    const thread: ThreadMessage[] = [
      { id: 'm2', subject: 'Something', from: 'carol@example.com', to: ['dave@example.com'], sentDate: new Date(), body: 'Hello', isFromCurrentUser: false }
    ];
    const id = setEmails(tp, thread);
  const mailbox = (global as any).Office.context.mailbox;
  mailbox.displayNewMessageForm = jest.fn();
  await (tp as any).forwardEmail(id);
  const arg = (mailbox.displayNewMessageForm as jest.Mock).mock.calls[0][0] as any;
  expect(/^FW:/i.test(arg.subject)).toBe(true);
  expect(arg.toRecipients.length).toBe(0);
  });
});

describe('AI auto-disable on health check failure', () => {
  beforeEach(() => {
    ensureDom();
    localStorage.clear();
    (global.fetch as jest.Mock).mockReset();
  });

  it('disables AI when health check fails', async () => {
    // Simulate health check failure
    (global.fetch as jest.Mock).mockResolvedValue({ ok: false, status: 500, statusText: 'err', text: async () => 'x' });
    const tp = new TaskpaneManager();
    // Provide config elements for loadConfiguration side effects
    (tp as any).configurationService.getConfiguration = jest.fn().mockResolvedValue(baseConfig);
    (tp as any).configurationService.saveConfiguration = jest.fn();
    await (tp as any).loadConfiguration();
    expect(localStorage.getItem('aiDisabled')).toBe('true');
  });
});
