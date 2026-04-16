import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import * as z from "zod/v4";

import type { AppConfig } from "./config.js";
import type { MicrosoftAuthService } from "./microsoft-auth.js";
import type { MicrosoftGraphClient } from "./microsoft-graph.js";

function pretty(value: unknown): string {
  return JSON.stringify(value, null, 2);
}

export function createMcpServer(services: {
  config: AppConfig;
  microsoftAuth: MicrosoftAuthService;
  graph: MicrosoftGraphClient;
}): McpServer {
  const server = new McpServer({
    name: "claude-m365-mcp",
    version: "0.1.0",
  });

  server.registerTool(
    "auth_status",
    {
      description:
        "Check whether the server is connected to Microsoft 365 and see any known delegated mailbox hints.",
      inputSchema: {},
      outputSchema: {
        connected: z.boolean(),
        account: z
          .object({
            name: z.string().nullable(),
            preferredUsername: z.string().nullable(),
            oid: z.string().nullable(),
            tid: z.string().nullable(),
          })
          .nullable(),
        expiresAt: z.number().nullable(),
        knownMailboxes: z.array(z.string()),
        localStatusUrl: z.string(),
        microsoftConnectUrl: z.string(),
        microsoftDisconnectUrl: z.string(),
      },
    },
    async () => {
      const status = await services.microsoftAuth.getStatus();
      const structuredContent = {
        connected: status.connected,
        account: status.account
          ? {
              name: status.account.name ?? null,
              preferredUsername: status.account.preferredUsername ?? null,
              oid: status.account.oid ?? null,
              tid: status.account.tid ?? null,
            }
          : null,
        expiresAt: status.expiresAt ?? null,
        knownMailboxes: status.knownMailboxes,
        localStatusUrl: services.config.localBaseUrl.toString(),
        microsoftConnectUrl: new URL(
          "/auth/microsoft/start",
          services.config.localBaseUrl,
        ).toString(),
        microsoftDisconnectUrl: new URL(
          "/auth/microsoft/disconnect",
          services.config.localBaseUrl,
        ).toString(),
      };

      return {
        content: [
          {
            type: "text",
            text: pretty(structuredContent),
          },
        ],
        structuredContent,
      };
    },
  );

  server.registerTool(
    "mail_list",
    {
      description:
        "List messages from a mailbox folder. Use mailbox for shared/delegated mailboxes the signed-in Microsoft user can access.",
      inputSchema: {
        mailbox: z.string().optional(),
        folder: z.string().default("Inbox"),
        top: z.number().int().min(1).max(100).default(25),
      },
      outputSchema: {
        mailbox: z.string(),
        folder: z.string(),
        messages: z.array(
          z.object({
            id: z.string(),
            subject: z.string(),
            from: z.string().nullable(),
            receivedDateTime: z.string().nullable(),
            sentDateTime: z.string().nullable(),
            bodyPreview: z.string(),
            webLink: z.string().nullable(),
            isDraft: z.boolean(),
            conversationId: z.string().nullable(),
          }),
        ),
      },
    },
    async ({ mailbox, folder, top }) => {
      const structuredContent = await services.graph.listMessages({ mailbox, folder, top });
      return {
        content: [{ type: "text", text: pretty(structuredContent) }],
        structuredContent,
      };
    },
  );

  server.registerTool(
    "mail_search",
    {
      description:
        "Search a mailbox using Microsoft Graph $search. Use mailbox for shared/delegated mailboxes.",
      inputSchema: {
        mailbox: z.string().optional(),
        query: z.string(),
        top: z.number().int().min(1).max(50).default(10),
      },
      outputSchema: {
        mailbox: z.string(),
        query: z.string(),
        messages: z.array(
          z.object({
            id: z.string(),
            subject: z.string(),
            from: z.string().nullable(),
            receivedDateTime: z.string().nullable(),
            sentDateTime: z.string().nullable(),
            bodyPreview: z.string(),
            webLink: z.string().nullable(),
            isDraft: z.boolean(),
            conversationId: z.string().nullable(),
          }),
        ),
      },
    },
    async ({ mailbox, query, top }) => {
      const structuredContent = await services.graph.searchMessages({ mailbox, query, top });
      return {
        content: [{ type: "text", text: pretty(structuredContent) }],
        structuredContent,
      };
    },
  );

  server.registerTool(
    "mail_get",
    {
      description:
        "Get the full details and body of one message by ID. Use mailbox for shared/delegated mailboxes.",
      inputSchema: {
        mailbox: z.string().optional(),
        messageId: z.string(),
      },
      outputSchema: {
        mailbox: z.string(),
        message: z.object({
          id: z.string(),
          subject: z.string(),
          from: z.string().nullable(),
          to: z.array(z.string()),
          cc: z.array(z.string()),
          bcc: z.array(z.string()),
          receivedDateTime: z.string().nullable(),
          sentDateTime: z.string().nullable(),
          bodyPreview: z.string(),
          body: z.object({
            contentType: z.string(),
            content: z.string(),
          }),
          webLink: z.string().nullable(),
          isDraft: z.boolean(),
          importance: z.string().nullable(),
          conversationId: z.string().nullable(),
        }),
      },
    },
    async ({ mailbox, messageId }) => {
      const structuredContent = await services.graph.getMessage({ mailbox, messageId });
      return {
        content: [{ type: "text", text: pretty(structuredContent) }],
        structuredContent,
      };
    },
  );

  server.registerTool(
    "mail_list_drafts",
    {
      description:
        "List draft messages from the default Drafts folder. Use mailbox for shared/delegated mailboxes.",
      inputSchema: {
        mailbox: z.string().optional(),
        top: z.number().int().min(1).max(100).default(25),
      },
      outputSchema: {
        mailbox: z.string(),
        drafts: z.array(
          z.object({
            id: z.string(),
            subject: z.string(),
            from: z.string().nullable(),
            receivedDateTime: z.string().nullable(),
            sentDateTime: z.string().nullable(),
            bodyPreview: z.string(),
            webLink: z.string().nullable(),
            isDraft: z.boolean(),
            conversationId: z.string().nullable(),
          }),
        ),
      },
    },
    async ({ mailbox, top }) => {
      const structuredContent = await services.graph.listDrafts({ mailbox, top });
      return {
        content: [{ type: "text", text: pretty(structuredContent) }],
        structuredContent,
      };
    },
  );

  server.registerTool(
    "mail_create_draft",
    {
      description:
        "Create a new draft email. Use mailbox for shared/delegated mailboxes and optionally set from when you need a specific sender.",
      inputSchema: {
        mailbox: z.string().optional(),
        subject: z.string(),
        to: z.array(z.string()).default([]),
        cc: z.array(z.string()).optional(),
        bcc: z.array(z.string()).optional(),
        body: z.string(),
        bodyType: z.enum(["text", "html"]).default("text"),
        from: z.string().optional(),
      },
      outputSchema: {
        mailbox: z.string(),
        draft: z.object({
          id: z.string(),
          subject: z.string(),
          from: z.string().nullable(),
          receivedDateTime: z.string().nullable(),
          sentDateTime: z.string().nullable(),
          bodyPreview: z.string(),
          webLink: z.string().nullable(),
          isDraft: z.boolean(),
          conversationId: z.string().nullable(),
        }),
      },
    },
    async ({ mailbox, subject, to, cc, bcc, body, bodyType, from }) => {
      const structuredContent = await services.graph.createDraft({
        mailbox,
        subject,
        to,
        cc,
        bcc,
        body,
        bodyType,
        from,
      });
      return {
        content: [{ type: "text", text: pretty(structuredContent) }],
        structuredContent,
      };
    },
  );

  server.registerTool(
    "mail_send_draft",
    {
      description:
        "Send an existing draft message by ID. Use mailbox for shared/delegated mailboxes.",
      inputSchema: {
        mailbox: z.string().optional(),
        messageId: z.string(),
      },
      outputSchema: {
        mailbox: z.string(),
        messageId: z.string(),
        sent: z.literal(true),
      },
    },
    async ({ mailbox, messageId }) => {
      const structuredContent = await services.graph.sendDraft({ mailbox, messageId });
      return {
        content: [{ type: "text", text: pretty(structuredContent) }],
        structuredContent,
      };
    },
  );

  server.registerTool(
    "mail_move",
    {
      description:
        "Move a message to another folder. Pass a well-known folder name like Archive or DeletedItems, or set destinationFolderIsId when passing a raw folder ID.",
      inputSchema: {
        mailbox: z.string().optional(),
        messageId: z.string(),
        destinationFolder: z.string(),
        destinationFolderIsId: z.boolean().default(false),
      },
      outputSchema: {
        mailbox: z.string(),
        destinationFolder: z.string(),
        movedMessage: z.object({
          id: z.string(),
          subject: z.string(),
          from: z.string().nullable(),
          receivedDateTime: z.string().nullable(),
          sentDateTime: z.string().nullable(),
          bodyPreview: z.string(),
          webLink: z.string().nullable(),
          isDraft: z.boolean(),
          conversationId: z.string().nullable(),
        }),
      },
    },
    async ({ mailbox, messageId, destinationFolder, destinationFolderIsId }) => {
      const structuredContent = await services.graph.moveMessage({
        mailbox,
        messageId,
        destinationFolder,
        destinationFolderIsId,
      });
      return {
        content: [{ type: "text", text: pretty(structuredContent) }],
        structuredContent,
      };
    },
  );

  server.registerTool(
    "calendar_list_events",
    {
      description:
        "List events in the default calendar over a time window. Use mailbox for shared/delegated calendars.",
      inputSchema: {
        mailbox: z.string().optional(),
        start: z.string().optional(),
        end: z.string().optional(),
        top: z.number().int().min(1).max(100).default(25),
      },
      outputSchema: {
        mailbox: z.string(),
        window: z.object({
          start: z.string(),
          end: z.string(),
        }),
        events: z.array(
          z.object({
            id: z.string(),
            subject: z.string(),
            webLink: z.string().nullable(),
            start: z.object({
              dateTime: z.string().nullable(),
              timeZone: z.string().nullable(),
            }),
            end: z.object({
              dateTime: z.string().nullable(),
              timeZone: z.string().nullable(),
            }),
            location: z.string().nullable(),
            attendees: z.array(
              z.object({
                address: z.string().nullable(),
                name: z.string().nullable(),
                type: z.string().nullable(),
                response: z.string().nullable(),
              }),
            ),
            bodyPreview: z.string(),
            body: z.object({
              contentType: z.string(),
              content: z.string(),
            }),
          }),
        ),
      },
    },
    async ({ mailbox, start, end, top }) => {
      const structuredContent = await services.graph.listEvents({ mailbox, start, end, top });
      return {
        content: [{ type: "text", text: pretty(structuredContent) }],
        structuredContent,
      };
    },
  );

  server.registerTool(
    "calendar_create_event",
    {
      description:
        "Create an event in the default calendar. Use mailbox for shared/delegated calendars.",
      inputSchema: {
        mailbox: z.string().optional(),
        subject: z.string(),
        start: z.string(),
        end: z.string(),
        timeZone: z.string().default("UTC"),
        attendees: z.array(z.string()).optional(),
        body: z.string().optional(),
        bodyType: z.enum(["text", "html"]).default("text"),
        location: z.string().optional(),
      },
      outputSchema: {
        mailbox: z.string(),
        event: z.object({
          id: z.string(),
          subject: z.string(),
          webLink: z.string().nullable(),
          start: z.object({
            dateTime: z.string().nullable(),
            timeZone: z.string().nullable(),
          }),
          end: z.object({
            dateTime: z.string().nullable(),
            timeZone: z.string().nullable(),
          }),
          location: z.string().nullable(),
          attendees: z.array(
            z.object({
              address: z.string().nullable(),
              name: z.string().nullable(),
              type: z.string().nullable(),
              response: z.string().nullable(),
            }),
          ),
          bodyPreview: z.string(),
          body: z.object({
            contentType: z.string(),
            content: z.string(),
          }),
        }),
      },
    },
    async ({ mailbox, subject, start, end, timeZone, attendees, body, bodyType, location }) => {
      const structuredContent = await services.graph.createEvent({
        mailbox,
        subject,
        start,
        end,
        timeZone,
        attendees,
        body,
        bodyType,
        location,
      });
      return {
        content: [{ type: "text", text: pretty(structuredContent) }],
        structuredContent,
      };
    },
  );

  return server;
}
