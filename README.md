MarkdownForOutlook
==================

Add-in for Microsoft Outlook to allow transforming email body text to Html using markdown syntax.

### Usage
Ctrl+Alt+M - will format/unformat the active email's body text

### Notes
See MarkdownSpecification.md for what is supported. Unfortunately not all the GitHub markdown features are supported.

### Todo
* recognize and show warning whenever undo text will clobber changes made by the user
* don't transform user signature and previous messages in the thread
* cleanup old mail items (will hooking into send work?)
