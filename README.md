# New-User
Generic create user script.

The goal of this script is to be able to create a new AD user and mailbox. Regardless of the version used in that environment.

The script will prompt the user for a few variables and credentials, also some environment details are requested when needed.
When Exchange is selected, the script will automatically detect the version & if possible the server.

The Office 365 part will try to create the new user, if enough licenses are available.

## Future 
Now that the working baseline is done, the next step is to transform the script and move different bits into functions that can be reused.

## Contributing
Create a fork of the project into your own reposity. Make all your necessary changes and create a pull request with a description on what was added or removed and details explaining the changes in lines of code. If approved, project owners will merge it.
 