# About

This is a reproduction repo for issue https://github.com/OfficeDev/office-js/issues/1088

# Steps to reproduce

## Preparations
1) Have OSX Catalina installed
2) Have up-to-date Outlook app installed
![Image of outlook version](outlookversion.png)
4) Send an email that has below text in it:
```
The wrong line:
M: 516-800-1000 / Fax: 516-900-2000
```

# Using the app
4) [Sideload manifest.xml](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing) to outlook app 
5) Install dependencies with `npm ci`
6) Start the app with `npm run dev-server`

# Expected outcome
When I have the two numbers highlighted and I click them, the contextual addin appears with the number i clicked.

# What actually happens
I have both numbers highlighted but the addin shows the first one even if i click the second.