# TOPlunch subscriber

[AutoIt](https://www.autoitscript.com/site/) script to ask colleagues whether they'll join lunch today so the proper amount of food can be prepared.

## Setup

Just make sure to run `TOPlunch subscriber.exe` during login of all colleagues.

## Daily usage

The user is presented with four options:

1. Yes - calls an http request to subscribe the user, opens a webpage to confirm this to the user and registers the user made a choice today.
2. No - registers the user made a choice today
3. Snooze - reminds the user in 15 minutes
4. Never ask me again - registers the user doesn't want to join lunch (this way) ever again

Choices are saved in `%userprofile%\TOPlunch subscriber.ini` to only ask the user once a day even after a reboot of Windows. This file can simply be removed as well if a user changes his/her mind about "Never ask me again".

Note: if it's past 11:30 the program doesn't do anything as it's already to late to take the user into account for food preparations. Because of this, Snooze is disabled after 11:15.