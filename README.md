# WordDocumentMarkupCleaner

**Caution! Changes your docx Files! Currupted files, missing xml or external data may lead to data loss**

![Screenshot](screenshot.jpg?raw=true "Screenshot")

## Features: 

- Cleans OpenXML Word Documents (.docx) from unnecessary markup. 
- Uses OpenXmlPowerTools by Eric White
- Offers batchprocessing and backups (.docx.bak)

### Default settings:

```csharp
SimplifyMarkupSettings settings = new SimplifyMarkupSettings
    {
        AcceptRevisions = true,
        //setting this to false reduces translation quality, but if true some documents have XML format errors when opening
        NormalizeXml = true,        // Merges Run's in a paragraph with similar formatting 
        RemoveBookmarks = true,
        RemoveComments = true,
        RemoveContentControls = true,
        RemoveEndAndFootNotes = true,
        RemoveFieldCodes = false, //true,
        RemoveGoBackBookmark = true,
        RemoveHyperlinks = false,
        RemoveLastRenderedPageBreak = true,
        RemoveMarkupForDocumentComparison = true,
        RemovePermissions = false,
        RemoveProof = true,
        RemoveRsidInfo = true,
        RemoveSmartTags = true,
        RemoveSoftHyphens = true,
        RemoveWebHidden = true,
        ReplaceTabsWithSpaces = false
    };
```