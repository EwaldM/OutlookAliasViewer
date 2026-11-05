# Outlook Alias Viewer

A small VBA macro for Microsoft Outlook (Classic) that displays the
original **To:** recipient address(es) from a message's transport
headers — useful when your mailbox has multiple alias addresses and
Outlook's To: field only shows your primary address.

## Purpose

When Exchange delivers a message to an alias, it rewrites the To:
field to your primary SMTP address. The original alias is preserved
only inside the raw `PR_TRANSPORT_MESSAGE_HEADERS` property. This
macro extracts and displays that address on demand.

## Limitations

- **External messages only.** `PR_TRANSPORT_MESSAGE_HEADERS` is
  populated for messages that arrived via the internet (MIME).
  Internal Exchange-to-Exchange messages typically do **not** contain
  this property, so the macro will show "None" for those.
- **Outlook Classic only.** Not compatible with the new Outlook
  (Windows) or Outlook on the web.
- **No column support.** This is a manual, on-demand viewer (MsgBox),
  not an automatic list-view column.
- **Macro security required.** You must enable macro notifications in
  the Trust Center.

## Installation

1. In Outlook, go to **File → Options → Customize Ribbon** and check
   **Developer** under Main Tabs. Click **OK**.
2. Press **Alt+F11** to open the VBA editor.
3. **Insert → Module**.
4. Paste the contents of `ShowAlias.bas`.
5. **Ctrl+S** to save. Close the VBA editor.
6. Go to **File → Options → Trust Center → Trust Center Settings…
   → Macro Settings** and select
   **"Notifications for all macros"**.
7. Restart Outlook.

### Add a ribbon button

1. **File → Options → Customize Ribbon**.
2. On the right, select the **Message** tab → **New Group…** →
   name it (e.g. "My Tools").
3. On the left, set **Choose commands from** → **Macros**.
4. Select **ShowAlias** → **Add** → **OK**.

Optionally right-click the new button to assign a shortcut key.

## Usage

1. Select a message in the list (or open it in a popup window).
2. Click the **ShowAlias** ribbon button (or use your shortcut).
3. A message box displays the original To: address(es).
   - Multiple addresses are comma-separated.
   - If no valid address is found, it shows "None".

## License

[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)