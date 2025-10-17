# HashFix - Outlook Email Hashtag Fixer

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Outlook](https://img.shields.io/badge/Outlook-Add--in-0078D4)
![License](https://img.shields.io/badge/license-MIT-green)

An Outlook add-in that automatically fixes improperly formatted hashtags in your emails before sending.

## Overview

HashFix is an event-based Outlook add-in that intercepts emails when you click "Send" and automatically corrects common hashtag formatting mistakes. It runs silently in the background and requires no manual intervention.

### What It Fixes

| Before (Incorrect) | After (Fixed) |
|-------------------|---------------|
| `# tag` | `#tag` |
| `#  marketing` | `#marketing` |
| `# 2024` | `#2024` |
| `#    team` | `#team` |

## Features

- **Automatic Correction**: Fixes hashtags without requiring manual intervention
- **OnSend Event**: Runs automatically when you click Send
- **Activity Tracking**: View statistics and history in the taskpane
- **HTML-Safe**: Preserves email formatting, links, and styles
- **Soft-Block Mode**: Allows sending even if processing fails (fail-safe)
- **Privacy-Focused**: All processing happens locally in your browser

## Installation

### Method 1: Outlook on the Web (Office 365)

#### New Outlook Web (Current UI)
1. Go to [Outlook on the web](https://outlook.office.com) and sign in
2. Click the **Apps** icon (waffle/grid icon) in the left sidebar
3. Click **Get Add-ins** (or search for it)
4. Select **My add-ins** tab on the left
5. Scroll down and click **➕ Add a custom add-in** → **Add from file**
6. Download `manifest.xml` from: https://David-Summers.github.io/hashfix/manifest.xml
7. Click **Browse** and select the downloaded `manifest.xml` file
8. Click **Upload** and accept the warning
9. The add-in will appear in the ribbon when composing emails

#### Classic Outlook Web (Older UI)
1. Go to [Outlook on the web](https://outlook.office.com)
2. Click the **Settings** gear icon (top right)
3. Search for "add-ins" or go to **View all Outlook settings** → **General** → **Manage add-ins**
4. Click **My add-ins** (left panel)
5. Under **Custom add-ins**, click **➕ Add a custom add-in** → **Add from file**
6. Download and upload the `manifest.xml` file
7. Accept the warning and click **Install**

### Method 2: New Outlook for Windows (Microsoft 365)

1. Open **New Outlook** (toggle in top-right corner if in classic Outlook)
2. Click **Apps** in the left navigation bar
3. Click **Get Add-ins**
4. Select **My add-ins** tab
5. Click **➕ Add a custom add-in** → **Add from file**
6. Download `manifest.xml` from the repository
7. Browse to the file and click **Upload**
8. Restart Outlook

### Method 3: Classic Outlook Desktop (Windows)

1. Open **Outlook Desktop**
2. Go to **Home** tab → **Get Add-ins** (or **Store**)
3. Click **My add-ins** in the left panel
4. Click **➕ Add a custom add-in** → **Add from file**
   - If you see "Add from URL", you can use: `https://David-Summers.github.io/hashfix/manifest.xml`
5. Browse to the downloaded `manifest.xml` file
6. Click **Install**
7. Restart Outlook

### Method 4: Outlook for Mac

1. Download the `manifest.xml` file from: https://David-Summers.github.io/hashfix/manifest.xml
2. Open **Finder** and press **Cmd+Shift+G** (Go to Folder)
3. Paste this path:
   ```
   ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef
   ```
4. Create the `wef` folder if it doesn't exist
5. Copy `manifest.xml` into the `wef` folder
6. Restart Outlook

### Method 5: Admin Deployment (For IT Administrators)

Administrators can deploy HashFix organization-wide via:

1. **Microsoft 365 Admin Center**:
   - Go to **Settings** → **Integrated apps** → **Upload custom apps**
   - Upload `manifest.xml`
   - Assign to users/groups

2. **Centralized Deployment**:
   - Admin Center → **Settings** → **Services & add-ins** → **Deploy Add-in**
   - Upload the manifest
   - Configure deployment settings

### Download Manifest File

If you need to download the manifest file manually:
```
https://David-Summers.github.io/hashfix/manifest.xml
```

Right-click and "Save as" or use curl/wget:
```bash
curl -O https://David-Summers.github.io/hashfix/manifest.xml
```

## Usage

### Automatic Mode (Default)

Simply compose and send emails normally. HashFix will automatically scan and fix hashtags when you click "Send". You don't need to do anything special.

### Viewing Activity

1. In Outlook, compose a new email
2. Click the **"HashFix"** button in the ribbon
3. The taskpane will open showing:
   - Current status
   - Total fixes applied
   - Session statistics
   - Recent activity log

### Example

**Before sending:**
```
Hey team, check out these updates:
# marketing campaign is live
# 2024 goals are updated
#  newproduct launching soon
```

**After HashFix processes (automatically):**
```
Hey team, check out these updates:
#marketing campaign is live
#2024 goals are updated
#newproduct launching soon
```

## Requirements

- **Outlook Version**: Requires Mailbox API 1.10 or higher
  - Outlook 2021+
  - Outlook on the Web
  - Microsoft 365 (Office 365)
- **Permissions**: ReadWriteMailbox (to read and modify email content)

## Technical Details

### Architecture

- **Event Type**: OnMessageSend with SoftBlock mode
- **Pattern Matching**: Uses regex `/#\s+(\w+)/g`
- **HTML Processing**: DOMParser for safe HTML manipulation
- **Storage**: SessionStorage for activity logs (local only)

### How It Works

1. User composes an email and clicks "Send"
2. OnMessageSend event fires
3. HashFix retrieves the email body (HTML format)
4. Parses HTML using DOMParser
5. Recursively processes text nodes
6. Fixes hashtag patterns using regex
7. Updates email body if changes were made
8. Logs activity to sessionStorage
9. Allows the email to send

### Files

```
hashfix/
├── manifest.xml           # Add-in configuration
├── functions.html         # OnSend event handler
├── functions.js          # Compatibility layer
├── taskpane.html         # Activity dashboard UI
├── support.html          # Full documentation
├── icon-16.png           # 16x16 icon
├── icon-32.png           # 32x32 icon
├── icon-64.png           # 64x64 icon
└── README.md             # This file
```

## Privacy & Security

HashFix processes your email content **locally in the browser**. No email content is sent to external servers.

### What HashFix Does:
- ✅ Reads email body when you click "Send"
- ✅ Modifies hashtag formatting
- ✅ Stores anonymous statistics in sessionStorage (local only)

### What HashFix Does NOT Do:
- ❌ Send email content to external servers
- ❌ Store email content
- ❌ Access your contacts or mailbox items
- ❌ Track personal information

## Troubleshooting

### Add-in not appearing in Outlook
- Ensure you're using a supported version of Outlook
- Try restarting Outlook
- Verify the manifest URL is correct
- Check your internet connection

### Hashtags not being fixed
- Open browser console (F12) to check for errors
- Ensure the add-in is enabled in Outlook settings
- Verify Office.js is loading correctly
- Try sending a test email with `# test`

### TaskPane shows "0" fixes
- SessionStorage might have been cleared
- Click "Refresh Activity" to reload
- Note: Statistics reset when you close Outlook

## Development

### Local Development Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/David-Summers/hashfix.git
   cd hashfix
   ```

2. Serve files locally (requires a web server):
   ```bash
   # Using Python
   python -m http.server 8000

   # Or using Node.js
   npx http-server -p 8000
   ```

3. Update manifest.xml URLs to point to localhost:
   ```xml
   <bt:Url id="Functions.Url" DefaultValue="http://localhost:8000/functions.html"/>
   ```

4. Sideload the manifest in Outlook

### Testing

Create test emails with these patterns:
- `# tag` (single space)
- `#  tag` (double space)
- `#    tag` (multiple spaces)
- `# 123` (numbers)
- Mixed content with multiple hashtags

### Modifying the Fix Logic

Edit `functions.html` at line 97 to customize the regex pattern:

```javascript
function fixHashtagText(text) {
  // Current: /#\s+(\w+)/g
  // Customize as needed
  return text.replace(/#\s+(\w+)/g, '#$1');
}
```

## Browser Console Debugging

Press **F12** in Outlook to open developer tools and view logs:

```javascript
// HashFix logs messages like:
console.log("HashFix: Hashtags automatically corrected");
console.error("Failed to get email body:", error);
```

## Limitations

- Does not work in Outlook Mobile (OnMessageSend not supported)
- Only processes outgoing emails (not incoming)
- Statistics reset when Outlook is closed
- Requires internet connection (files hosted on GitHub Pages)

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly in Outlook
5. Submit a pull request

## Support

For help, bug reports, or feature requests:

- **GitHub Issues**: [github.com/David-Summers/hashfix/issues](https://github.com/David-Summers/hashfix/issues)
- **Documentation**: [support.html](https://David-Summers.github.io/hashfix/support.html)
- **Author**: Mark Robinson

## Version History

### 1.0.0 (2025)
- Initial release
- Automatic hashtag spacing correction
- OnMessageSend event handler
- Activity tracking and statistics
- Taskpane UI with dashboard

## License

MIT License - See LICENSE file for details

## Acknowledgments

- Built with [Office.js](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)
- Hosted on [GitHub Pages](https://pages.github.com/)
- Designed for Microsoft Outlook

---

**Made with care by Mark Robinson**

For more information, visit the [GitHub repository](https://github.com/David-Summers/hashfix).
