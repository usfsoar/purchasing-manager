# SOAR Purchasing Database Manager
This is the Google Apps Script project for SOAR's purchasing database. This 
code handles tracking of purchasing items, using our Google Sheets database
as a backend.

## Contributing
This code will not run as-is; you need some channel webhooks to put into the
`secret.template` file, before renaming it to `secret.js`. Also, the `clasp`
config file is not included, so you cannot push directly to the Google Apps
Script project. If you want to submit changes to this code and see them live
in SOAR's own database, submit a pull request and the code will be pushed to
Google Apps Script if merged.

If you want to use this code in your own project, you are free to under the GPL
license.

For more info contact [Ian Sanders](github.com/iansan5653).