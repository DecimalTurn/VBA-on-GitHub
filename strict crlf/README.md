So you've found the CRLF strict section. Good. Very Good. That means you are a rebel. A non-comforist.

The .gitattributes in this sub-folder is there to help you go against the desires of Git to convert you line endings to LF and your Windows "ANSI" encoding to UTF-8.

- Make sure to also add the .gitconfig file at the root of your repository
- Note that we can't include the wd file encoding otherwise Git will perform the conversion. We can't let Git know anything. The more it knows, the easier it is for Git to try and "help us" by converting our code.