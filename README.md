# WordDocumentMarkupCleaner
WordDocumentMarkupCleaner

[![Package Quality](http://npm.packagequality.com/shield/WordDocumentMarkupCleaner.svg)](http://packagequality.com/#?package=WordDocumentMarkupCleaner)
[![Inline docs](http://inch-ci.org/github/HansHammel/WordDocumentMarkupCleaner.svg?branch=master)](http://inch-ci.org/github/HansHammel/WordDocumentMarkupCleaner)
[![star this repo](http://githubbadges.com/star.svg?user=HansHammel&repo=WordDocumentMarkupCleaner&style=flat&color=fff&background=007ec6)](https://github.com/HansHammel/WordDocumentMarkupCleaner)
[![fork this repo](http://githubbadges.com/fork.svg?user=HansHammel&repo=WordDocumentMarkupCleaner&style=flat&color=fff&background=007ec6)](https://github.com/HansHammel/WordDocumentMarkupCleaner/fork)
[![david dependency](https://img.shields.io/david/HansHammel/WordDocumentMarkupCleaner.svg)](https://david-dm.org/HansHammel/WordDocumentMarkupCleaner)
[![david devDependency](https://img.shields.io/david/dev/HansHammel/WordDocumentMarkupCleaner.svg)](https://david-dm.org/HansHammel/WordDocumentMarkupCleaner)
[![david optionalDependency](https://img.shields.io/david/optional/HansHammel/WordDocumentMarkupCleaner.svg)](https://david-dm.org/HansHammel/WordDocumentMarkupCleaner)
[![david peerDependency](https://img.shields.io/david/peer/HansHammel/WordDocumentMarkupCleaner.svg)](https://david-dm.org/HansHammel/WordDocumentMarkupCleaner)
[![Known Vulnerabilities](https://snyk.io/test/github/HansHammel/WordDocumentMarkupCleaner/badge.svg)](https://snyk.io/test/github/HansHammel/WordDocumentMarkupCleaner) 
[![Greenkeeper badge](https://badges.greenkeeper.io/HansHammel/WordDocumentMarkupCleaner.svg)](https://greenkeeper.io/)

![Screenshot](screenshot.jpg?raw=true "Screenshot")


*** Caution! Changes your docx Files! Currupted files, missing xml or external data may lead to data loss***

Features: 

- Cleans OpenXML Word Documents (.docx) from unnecessary markup. 
- Uses OpenXmlPowerTools by Eric White
- Offers batchprocessing and backups (.docx.bak)

Default settings:

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