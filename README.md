# Dynamic images in Docs

# Overview
This is Google Apps Script for inserting dynamic images into Google Docs. Images are stored in Google Drive.

# Apps Script dependencies
None. 

# Docs Template
<div align="center"><img src="https://github.com/user-attachments/assets/de15b41c-0017-4bcb-8afb-87f0c774b735" /></div>

# Sheets
Check the boxes then click **Custom Menu -> Create documents**

<div align="center"><img src="https://github.com/user-attachments/assets/38499891-1967-4193-99ec-36de1c12feee" /></div>

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

