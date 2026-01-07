# Email Signature

## What

Automatically append an email signature when sending mail via the MCP server, if a signature file exists.

## Why

Email signatures are standard practice for professional communication. Users shouldn't have to manually include their signature in every email body.

## Requirements

- [x] If `~/.config/ops-personal-m365-mcp/signature.html` exists, use it
- [x] `useSignature` parameter on `mail_send` (default: true) to opt-out
- [x] `mail_reply` tool with `useSignature` defaulting to false
- [x] Change `isHtml` default to `true` (HTML is now the default format)

## Approach

Check for signature file existence in `sendMail`. If present and `useSignature` is true, append it to the email body. No additional tools needed - users manage the file manually or via handbook.

## Tasks

- [x] Create `src/utils/signature.ts` with `getSignature()` function
- [x] Change `isHtml` default from `false` to `true` in `mail.ts`
- [x] Add `useSignature` parameter to `sendMailSchema` in `mail.ts`
- [x] Modify `sendMail` to append signature when enabled and file exists
- [x] Update `mail_send` tool definition in `index.ts`

## Storage

```
~/.config/ops-personal-m365-mcp/signature.html
```

## Notes

- Signature is stored as HTML, managed manually by user
- Appended with `<br><br>` separator before signature
- Handbook integration: see `handbook-signature-generator-prompt.md`
