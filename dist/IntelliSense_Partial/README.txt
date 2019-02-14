This directory contains copies of host-specific IntelliSense files (Word, Excel, etc -- corresponding to the postosf project hosts), as well as the core OfficeRuntime and OfficeCore files, that get combined into "office-vsdoc.js" and "office.d.ts" at the root of the folder (alongside the other Office.js files).

The generation of these files is handled as part of a larger "otools\inc\osfclient\publish-tools\update-warehouse.bat" flow.

The partial files are copied as is from the tenant (and in the case of the d.ts file, are cleaned up in the process as well).  However, substitution of the word "(PREVIEW)" only happens at the very end, as part of "update-warehouse.bat", on the final combined "office-vsdoc.js" and "office.d.ts", but not on the partial files.

The generation of the "(PREVIEW)" text is driven off of "METADATA.json".