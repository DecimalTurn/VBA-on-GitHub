# VBA-Template
This project is a template you can use to start your next VBA project on GitHub.

## .gitignore

This file is there to help Git determine what files it should ignore when commiting changes to the project. Unlike for [VScode](https://github.com/github/gitignore/blob/main/Global/VisualStudioCode.gitignore), the VBE doesn't really create local or temporary files that you don't want to share outside of your computer. However, the Office application you are using might create temporary files that remain on the disk as long as the Office document is open to indicate to others that someone is working in the file ([more details](https://superuser.com/questions/405257/what-type-of-file-is-file)).

This is why an exclusion exists for this type of file:
```
~$*
```

Sometimes, you'll be working from an actual Office document and only want to share the code as .bas, .frm and .frx files. In that case, feel free to add extensions for Office documents extensions to your .gitignore file. Some of them are already in the template .gitignore file as commented out line. Just remove the "#" to add them back as needed.

### Weird brackets?
You'll notice that the .gitignore (and .gitattributes) make use of this syntax with brackets to specify file paths:
eg.
```ignore
*.[xX][lL][aA]
```

This is because git configurations are case-sensitive: if we don't want to use 2 lines for each file extension (uppercase and lowercase version), we need to take advantage of the [glob pattern syntax](https://en.wikipedia.org/wiki/Glob_(programming)#Syntax) to combine lowercase and uppercase letters into pairs contained in those bracket. It's a bit harder to read, but I'm sure you'll get used to it. So far, I've resisted the temptation to combine different file extensions that differ by only one letter, but I might merge those rows later.

## .gitattributes

The standard .gitattributes file in this template is probably longer that you need, but it's there to cover the vast majority of cases related to VBA developement. This version will make sure that VBA related files are encoded using UTF-8 and Unix-style line endings (LF) when they are committed to the Git index. This is to better work with the GitHub UI which is often expecting those specific encodings to work properly.

Disclaimer: Converting your files to use UTF-8 and LF line endings means that people that want to download your code might have a problem if they try to download the a raw file with code. Be ready to get some issues like [this one](https://github.com/VBA-tools/VBA-Dictionary/issues/38) if you project becomes really popular. 
![Alt text](./docs/img/ScreenCapDownloadRawFile.png)

So far, this download issue is the only known issue related to encoding with the standard approach. However, if you're feeling more rebelious. If you consider that using LF inside .bas files is a blasphemous act, you could switch to use the .gitattributes file located inside the the `strict crlf` folder.



## To add

- Regarding the attribute `linguist-language=vba`, I choose not to include it because I believe that GitHub's system to detect if a file is VBA or not is decent and if your files aren't detected as VBA, it might be telling you somethig about your code.
  - Make sure that your files include the VBE's metadata such as the `Attribute VB_Name = "..."`
  - Use the right file extensions
  - If you .bas files doesn't contain any VBA specific syntax, maybe it's OK if is classified as VB6 instead.
- Mention git lfs?
