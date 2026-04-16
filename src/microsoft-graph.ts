import type { MicrosoftAuthService } from "./microsoft-auth.js";

type GraphEmailAddress = {
  address?: string;
  name?: string;
};

type GraphRecipient = {
  emailAddress?: GraphEmailAddress;
};

type GraphMessage = {
  id: string;
  subject?: string;
  from?: GraphRecipient;
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  bccRecipients?: GraphRecipient[];
  receivedDateTime?: string;
  sentDateTime?: string;
  bodyPreview?: string;
  body?: {
    contentType?: string;
    content?: string;
  };
  webLink?: string;
  isDraft?: boolean;
  importance?: string;
  conversationId?: string;
};

type GraphMailFolder = {
  id: string;
  displayName?: string;
};

type GraphEvent = {
  id: string;
  subject?: string;
  webLink?: string;
  start?: {
    dateTime?: string;
    timeZone?: string;
  };
  end?: {
    dateTime?: string;
    timeZone?: string;
  };
  location?: {
    displayName?: string;
  };
  attendees?: Array<{
    emailAddress?: GraphEmailAddress;
    type?: string;
    status?: {
      response?: string;
      time?: string;
    };
  }>;
  bodyPreview?: string;
  body?: {
    contentType?: string;
    content?: string;
  };
};

export type MessageSummary = {
  id: string;
  subject: string;
  from: string | null;
  receivedDateTime: string | null;
  sentDateTime: string | null;
  bodyPreview: string;
  webLink: string | null;
  isDraft: boolean;
  conversationId: string | null;
};

export type FullMessage = {
  id: string;
  subject: string;
  from: string | null;
  to: string[];
  cc: string[];
  bcc: string[];
  receivedDateTime: string | null;
  sentDateTime: string | null;
  bodyPreview: string;
  body: {
    contentType: string;
    content: string;
  };
  webLink: string | null;
  isDraft: boolean;
  importance: string | null;
  conversationId: string | null;
};

export type CalendarEvent = {
  id: string;
  subject: string;
  webLink: string | null;
  start: {
    dateTime: string | null;
    timeZone: string | null;
  };
  end: {
    dateTime: string | null;
    timeZone: string | null;
  };
  location: string | null;
  attendees: Array<{
    address: string | null;
    name: string | null;
    type: string | null;
    response: string | null;
  }>;
  bodyPreview: string;
  body: {
    contentType: string;
    content: string;
  };
};

function mapEmailAddress(recipient: GraphRecipient | undefined): string | null {
  return recipient?.emailAddress?.address ?? null;
}

function mapRecipients(recipients: GraphRecipient[] | undefined): string[] {
  return (recipients ?? [])
    .map((recipient) => recipient.emailAddress?.address)
    .filter((value): value is string => Boolean(value));
}

export class MicrosoftGraphClient {
  constructor(private readonly authService: MicrosoftAuthService) {}

  async listMessages(input: {
    mailbox?: string;
    folder?: string;
    top?: number;
  }): Promise<{
    mailbox: string;
    folder: string;
    messages: MessageSummary[];
  }> {
    const mailbox = input.mailbox ?? "me";
    const folder = input.folder ?? "Inbox";
    const top = Math.min(input.top ?? 25, 100);
    const base = this.basePath(input.mailbox);
    const query = new URLSearchParams({
      $top: String(top),
      $select:
        "id,subject,from,receivedDateTime,sentDateTime,bodyPreview,webLink,isDraft,conversationId",
    });
    const result = await this.request<{ value: GraphMessage[] }>(
      `${base}/mailFolders('${encodeURIComponent(folder)}')/messages?${query.toString()}`,
    );

    return {
      mailbox,
      folder,
      messages: result.value.map((message) => this.mapMessageSummary(message)),
    };
  }

  async searchMessages(input: {
    mailbox?: string;
    query: string;
    top?: number;
  }): Promise<{
    mailbox: string;
    query: string;
    messages: MessageSummary[];
  }> {
    const mailbox = input.mailbox ?? "me";
    const top = Math.min(input.top ?? 10, 50);
    const base = this.basePath(input.mailbox);
    const query = new URLSearchParams({
      $top: String(top),
      $search: `"${input.query.replaceAll('"', '\\"')}"`,
      $select:
        "id,subject,from,receivedDateTime,sentDateTime,bodyPreview,webLink,isDraft,conversationId",
    });
    const result = await this.request<{ value: GraphMessage[] }>(
      `${base}/messages?${query.toString()}`,
      {
        headers: {
          ConsistencyLevel: "eventual",
        },
      },
    );

    return {
      mailbox,
      query: input.query,
      messages: result.value.map((message) => this.mapMessageSummary(message)),
    };
  }

  async getMessage(input: {
    mailbox?: string;
    messageId: string;
  }): Promise<{
    mailbox: string;
    message: FullMessage;
  }> {
    const mailbox = input.mailbox ?? "me";
    const base = this.basePath(input.mailbox);
    const query = new URLSearchParams({
      $select:
        "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,bodyPreview,body,webLink,isDraft,importance,conversationId",
    });
    const message = await this.request<GraphMessage>(
      `${base}/messages/${encodeURIComponent(input.messageId)}?${query.toString()}`,
    );

    return {
      mailbox,
      message: this.mapFullMessage(message),
    };
  }

  async listDrafts(input: {
    mailbox?: string;
    top?: number;
  }): Promise<{
    mailbox: string;
    drafts: MessageSummary[];
  }> {
    const mailbox = input.mailbox ?? "me";
    const base = this.basePath(input.mailbox);
    const top = Math.min(input.top ?? 25, 100);
    const query = new URLSearchParams({
      $top: String(top),
      $select:
        "id,subject,from,receivedDateTime,sentDateTime,bodyPreview,webLink,isDraft,conversationId",
    });
    const result = await this.request<{ value: GraphMessage[] }>(
      `${base}/mailFolders('Drafts')/messages?${query.toString()}`,
    );

    return {
      mailbox,
      drafts: result.value.map((message) => this.mapMessageSummary(message)),
    };
  }

  async createDraft(input: {
    mailbox?: string;
    subject: string;
    to: string[];
    cc?: string[];
    bcc?: string[];
    body: string;
    bodyType?: "text" | "html";
    from?: string;
  }): Promise<{
    mailbox: string;
    draft: MessageSummary;
  }> {
    const mailbox = input.mailbox ?? "me";
    const base = this.basePath(input.mailbox);
    const message = await this.request<GraphMessage>(`${base}/messages`, {
      method: "POST",
      body: JSON.stringify({
        subject: input.subject,
        body: {
          contentType: input.bodyType ?? "text",
          content: input.body,
        },
        toRecipients: this.toRecipients(input.to),
        ccRecipients: this.toRecipients(input.cc),
        bccRecipients: this.toRecipients(input.bcc),
        from: input.from
          ? {
              emailAddress: {
                address: input.from,
              },
            }
          : input.mailbox
            ? {
                emailAddress: {
                  address: input.mailbox,
                },
              }
            : undefined,
      }),
    });

    return {
      mailbox,
      draft: this.mapMessageSummary(message),
    };
  }

  async sendDraft(input: {
    mailbox?: string;
    messageId: string;
  }): Promise<{
    mailbox: string;
    messageId: string;
    sent: true;
  }> {
    const mailbox = input.mailbox ?? "me";
    const base = this.basePath(input.mailbox);
    await this.request<void>(`${base}/messages/${encodeURIComponent(input.messageId)}/send`, {
      method: "POST",
    });

    return {
      mailbox,
      messageId: input.messageId,
      sent: true,
    };
  }

  async moveMessage(input: {
    mailbox?: string;
    messageId: string;
    destinationFolder: string;
    destinationFolderIsId?: boolean;
  }): Promise<{
    mailbox: string;
    movedMessage: MessageSummary;
    destinationFolder: string;
  }> {
    const mailbox = input.mailbox ?? "me";
    const base = this.basePath(input.mailbox);
    const destinationId = input.destinationFolderIsId
      ? input.destinationFolder
      : await this.resolveFolderId(base, input.destinationFolder);
    const moved = await this.request<GraphMessage>(
      `${base}/messages/${encodeURIComponent(input.messageId)}/move`,
      {
        method: "POST",
        body: JSON.stringify({
          destinationId,
        }),
      },
    );

    return {
      mailbox,
      movedMessage: this.mapMessageSummary(moved),
      destinationFolder: input.destinationFolder,
    };
  }

  async listEvents(input: {
    mailbox?: string;
    start?: string;
    end?: string;
    top?: number;
  }): Promise<{
    mailbox: string;
    window: {
      start: string;
      end: string;
    };
    events: CalendarEvent[];
  }> {
    const mailbox = input.mailbox ?? "me";
    const start = input.start ?? new Date().toISOString();
    const end =
      input.end ??
      new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
    const top = Math.min(input.top ?? 25, 100);
    const base = this.basePath(input.mailbox);
    const query = new URLSearchParams({
      startDateTime: start,
      endDateTime: end,
      $top: String(top),
      $orderby: "start/dateTime",
      $select:
        "id,subject,webLink,start,end,location,attendees,bodyPreview,body",
    });
    const result = await this.request<{ value: GraphEvent[] }>(
      `${base}/calendarView?${query.toString()}`,
    );

    return {
      mailbox,
      window: { start, end },
      events: result.value.map((event) => this.mapEvent(event)),
    };
  }

  async createEvent(input: {
    mailbox?: string;
    subject: string;
    start: string;
    end: string;
    timeZone?: string;
    attendees?: string[];
    body?: string;
    bodyType?: "text" | "html";
    location?: string;
  }): Promise<{
    mailbox: string;
    event: CalendarEvent;
  }> {
    const mailbox = input.mailbox ?? "me";
    const path = input.mailbox
      ? `/users/${encodeURIComponent(input.mailbox)}/calendar/events`
      : "/me/calendar/events";
    const event = await this.request<GraphEvent>(path, {
      method: "POST",
      body: JSON.stringify({
        subject: input.subject,
        start: {
          dateTime: input.start,
          timeZone: input.timeZone ?? "UTC",
        },
        end: {
          dateTime: input.end,
          timeZone: input.timeZone ?? "UTC",
        },
        attendees: (input.attendees ?? []).map((address) => ({
          emailAddress: { address },
          type: "required",
        })),
        body: input.body
          ? {
              contentType: input.bodyType ?? "text",
              content: input.body,
            }
          : undefined,
        location: input.location
          ? {
              displayName: input.location,
            }
          : undefined,
      }),
    });

    return {
      mailbox,
      event: this.mapEvent(event),
    };
  }

  private async resolveFolderId(base: string, folderName: string): Promise<string> {
    const folder = await this.request<GraphMailFolder>(
      `${base}/mailFolders('${encodeURIComponent(folderName)}')?$select=id,displayName`,
    );

    return folder.id;
  }

  private basePath(mailbox?: string): string {
    return mailbox ? `/users/${encodeURIComponent(mailbox)}` : "/me";
  }

  private async request<T>(
    path: string,
    init?: RequestInit & {
      headers?: Record<string, string>;
    },
  ): Promise<T> {
    const accessToken = await this.authService.getAccessToken();
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      ...init,
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${accessToken}`,
        ...(init?.body ? { "Content-Type": "application/json" } : {}),
        ...(init?.headers ?? {}),
      },
    });

    if (response.status === 204) {
      return undefined as T;
    }

    const text = await response.text();
    const data = text ? JSON.parse(text) : undefined;

    if (!response.ok) {
      const errorMessage =
        data?.error?.message ??
        data?.error_description ??
        response.statusText;
      const errorCode = data?.error?.code;
      const detail = errorCode ? `${errorCode}: ${errorMessage}` : errorMessage;
      throw new Error(`Microsoft Graph request failed (${response.status}): ${detail}`);
    }

    return data as T;
  }

  private toRecipients(addresses?: string[]): GraphRecipient[] | undefined {
    const cleaned = (addresses ?? []).filter(Boolean);
    if (cleaned.length === 0) {
      return undefined;
    }

    return cleaned.map((address) => ({
      emailAddress: {
        address,
      },
    }));
  }

  private mapMessageSummary(message: GraphMessage): MessageSummary {
    return {
      id: message.id,
      subject: message.subject ?? "",
      from: mapEmailAddress(message.from),
      receivedDateTime: message.receivedDateTime ?? null,
      sentDateTime: message.sentDateTime ?? null,
      bodyPreview: message.bodyPreview ?? "",
      webLink: message.webLink ?? null,
      isDraft: message.isDraft ?? false,
      conversationId: message.conversationId ?? null,
    };
  }

  private mapFullMessage(message: GraphMessage): FullMessage {
    return {
      id: message.id,
      subject: message.subject ?? "",
      from: mapEmailAddress(message.from),
      to: mapRecipients(message.toRecipients),
      cc: mapRecipients(message.ccRecipients),
      bcc: mapRecipients(message.bccRecipients),
      receivedDateTime: message.receivedDateTime ?? null,
      sentDateTime: message.sentDateTime ?? null,
      bodyPreview: message.bodyPreview ?? "",
      body: {
        contentType: message.body?.contentType ?? "text",
        content: message.body?.content ?? "",
      },
      webLink: message.webLink ?? null,
      isDraft: message.isDraft ?? false,
      importance: message.importance ?? null,
      conversationId: message.conversationId ?? null,
    };
  }

  private mapEvent(event: GraphEvent): CalendarEvent {
    return {
      id: event.id,
      subject: event.subject ?? "",
      webLink: event.webLink ?? null,
      start: {
        dateTime: event.start?.dateTime ?? null,
        timeZone: event.start?.timeZone ?? null,
      },
      end: {
        dateTime: event.end?.dateTime ?? null,
        timeZone: event.end?.timeZone ?? null,
      },
      location: event.location?.displayName ?? null,
      attendees: (event.attendees ?? []).map((attendee) => ({
        address: attendee.emailAddress?.address ?? null,
        name: attendee.emailAddress?.name ?? null,
        type: attendee.type ?? null,
        response: attendee.status?.response ?? null,
      })),
      bodyPreview: event.bodyPreview ?? "",
      body: {
        contentType: event.body?.contentType ?? "text",
        content: event.body?.content ?? "",
      },
    };
  }
}
