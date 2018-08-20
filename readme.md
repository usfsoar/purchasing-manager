# SOAR Purchasing Database Manager
This is the Google Apps Script project for SOAR's purchasing database. This 
code handles tracking of purchasing items, using our Google Sheets database
as a backend.

## Contributing
If you're a SOAR member who wants to contribute to this project, first request
that a financial officer gives you editing access to the [purchasing database](https://docs.google.com/spreadsheets/d/14Q5FTNgsqDfZEpBl9nmuZF7c3tK5NTO0oAl_P9269t8/edit#gid=0)
and this repository.

Then you'll need to install [Git](https://git-scm.com/downloads) and [Node, with NPM](https://nodejs.org/en/download/).
Open a command prompt in whatever folder you'd like to put the project's code. Run
`npm install @google/clasp -g`. After installing clasp, run `git clone https://github.com/usfsoar/purchasing-manager.git`.
Enter the folder with `cd purchasing-manager` and run `clasp login`. Open the link
provided and log in to clasp, allowing it the permissions it requests. Then run
`clasp pull` to fetch the *secret.gs* file from Google Apps Scripts. Now you have
all the files you need to edit the project.

Edit the code. With every major group of changes, run `git add .` and 
`git commit -m "Description of your changes"` and then `git push`. Finally, run 
`clasp push` when you're done, in order to push the new code live to the Google
Apps Script project. 

Having trouble? Message *@Ian Sanders* on Slack.