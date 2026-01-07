# Handbook Signature Generator - Claude Integration

## Task

Update the existing email signature generator in the Awolve handbook to support copying signatures for use with Claude Code's M365 MCP server.

## Requirements

### 1. Copy Button for Claude

Add a "Copy for Claude" button next to the existing copy functionality that:
- Copies the generated HTML signature to clipboard
- Shows a success toast/notification

### 2. Instructions Section

Add a collapsible "Using with Claude" section below the generator with these instructions:

```markdown
## Using Your Signature with Claude

After copying your signature, save it to:

```
~/.config/ops-personal-m365-mcp/signature.html
```

When sending emails through Claude, your signature is automatically appended.

To send without a signature, just say:
- "Send email to X without a signature"
```

### 3. UI Considerations

- Keep the existing generator functionality unchanged
- Add the Claude-specific button in a way that doesn't clutter the existing UI
- Consider adding a small Claude logo/icon to the button
- Instructions should be collapsed by default to not overwhelm users who don't use Claude

## Technical Notes

- The signature HTML should be copied exactly as generated (no modifications)
- Signature is stored in `~/.config/ops-personal-m365-mcp/signature.html`
- Users can also manually edit this file directly if needed
- The MCP server handles all signature storage - the handbook just needs to facilitate copying
