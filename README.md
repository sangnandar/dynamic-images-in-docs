# Dynamic images in Docs

# Overview
This is Google Apps Script for inserting dynamic images into Google Docs. Images are stored in Google Drive.

# Apps Script dependencies
None. 

# Docs Template
(template screenshot)

# Sheets
Check the boxes then click **Custom Menu -> Create documents**
(sheets screenshot)

# Sheets dependencies
**DO NOT** re-arrange columns or delete columns affected by the ranges.

Read
```
'Sheet1'!A2:B
```
Write
```
None
```

# Folder structure
```
imageFolder/
├── John Smith/
│   ├── Figure 1.png
│   └── Figure 2.jpg
├── Jane Doe/
│   ├── Figure 1.jpg
│   └── Figure 2.png
└── (etc)

docsFolder/
├── John Smith.gdoc
├── Jane Doe.gdoc
└── (etc)
```

