# VBA-Template
This project is a template you can use to start your next VBA project on GitHub.

## .gitignore

This file is there to help Git determine what files it should ignore when commiting changes to the project. Unlike for [VScode](https://github.com/github/gitignore/blob/main/Global/VisualStudioCode.gitignore), the VBE doesn't really create local or temporary files that you don't want to share outside of your computer. However, the Office application you are using might create temporary files that remain on the disk as long as the Office document is open to indicate to others that someone is working in the file ([more details](https://superuser.com/questions/405257/what-type-of-file-is-file)).

This is why an exclusion exists for this type of file:
```
~$*
```

Sometimes, you'll be working from an actual Office document and only want to share the code as .bas, .frm and .frx files. In that case, feel free to add extensions for Office documents extensions to your .gitignore file. Some of them are already in the template .gitignore file as commented out line. Just remove the "#" to add them back as needed.
