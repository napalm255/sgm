###SGM (Special Group Monitor)

 Open "SGM.exe.config" and edit the databaseConnectionString replacing "SERVERNAME", "DATABASE_NAME", "DATABASE_USER" and "DATABASE_PASSWORD"

 To install SGM, from the command line, run "installutil sgm.exe"

 To uninstall SGM, first stop the SGM service, then from the command line, run "installutil /u sgm.exe"

 The service leverages a SQL database for it's settings and logging. Within the database you define what groups/OUs to monitor or exclude from monitoring, who to alert, etc.

