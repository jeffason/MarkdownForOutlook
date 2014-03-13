MarkdownForOutlook
==================

Add-in for Microsoft Outlook to allow transforming email body text to Html using markdown syntax.

Ctrl+Alt+M - will format/unformat the active email's body text

See MarkdownSpecification.md for what is supported. Unfortunately not all the GitHub markdown features are supported.

Some known TODOS include:
* recognize and show warning whenever undo text will clobber changes made by the user
* don't transform user signature and previous messages in the thread
* cleanup old mail items (will hooking into send work?)
