Assumes use of the 'Terminal' application.

## Retrieving documents from archive-site

- Execute script:  ```pwsh -NoProfile -File .\Retrieve-DocumentsFromArchiveSite.ps1 | Tee-Object -FilePath .\Retrieve-DocumentsFromArchiveSite.out```

## Extracting raw-text from PDF-files in archive folder

- Execute script:  ```pwsh -NoProfile -File .\Extract-RawTextDocsFromArchive.ps1 | Tee-Object -FilePath .\Extract-RawTextDocsFromArchive.out```

## Joining raw-text docs into JSON dictionary

- Execute script:  ```pwsh -NoProfile -File .\Join-RawTextDocsToJsonDictionary.ps1 | Tee-Object -FilePath .\Join-RawTextDocsToJsonDictionary.out```
