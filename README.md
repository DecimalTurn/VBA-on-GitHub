# Using VBA on GitHub
This project is here to help you configure your VBA project on GitHub and discuss related issues.

# Content

- [Why should you include `.gitignore` and `.gitattributes` files to your project?](#why-should-you-include-gitignore-and-gitattributes-files-to-your-project)
- [.gitignore](#gitignore)
- [.gitattributes](#gitattributes)
- [Should you include a `.editorconfig` file in your repository?](#should-you-include-a-editorconfig-file-in-your-repository)
- [.editorconfig](#editorconfig)

# Why should you include .gitignore and .gitattributes files to your project?

**.gitignore**: This file is there to help Git determine what files it should ignore when committing changes to the project. It will allow you to avoid uploading content that you don't want to.

**.gitattributes**: This file will help you make sure that there won't be errors due to Git aplying conversions of your files.

# .gitignore

A template .gitignore file is provided here: 

📖 [Template](https://github.com/DecimalTurn/VBA-on-GitHub/blob/main/gitignore/.gitignore)

## Explanations

Luckily for us, the [VBE][VBEDEF] doesn't really create local or temporary files that you don't want to share outside of your computer (unlike for other IDEs like [VS Code](https://github.com/github/gitignore/blob/main/Global/VisualStudioCode.gitignore)). However, the Office application you are using might create temporary files that remain on the disk as long as the Office document is open to indicate to others that someone is working in the file ([more details](https://superuser.com/questions/405257/what-type-of-file-is-file)).

This is why an exclusion exists for this type of file (starting with ~$):
```
~$*
```

Sometimes, you'll be working from an actual Office document and only want to share the code as `.bas`, `.frm` and `.frx` files (not the office document itself). In that case, feel free to add extensions for Office documents extensions to your .gitignore file. Some of them are already in the template .gitignore file as commented-out lines. Just remove the "#" at the start of the line to add them back as needed.

### Weird brackets?
You'll notice that the .gitignore (and .gitattributes) files make use of this syntax with brackets to specify file paths:
eg.
```ignore
*.[xX][lL][aA]
```

This is because git configurations are case-sensitive: if we don't want to use 2 lines for each file extension (uppercase and lowercase version), we need to take advantage of the [glob pattern syntax](https://en.wikipedia.org/wiki/Glob_(programming)#Syntax) to combine lowercase and uppercase letters into pairs contained in those brackets. It's a bit harder to read, but I'm sure you'll get used to it. 

# .gitattributes

A template .gitattributes file is provided here: 

📖 [Template](https://github.com/DecimalTurn/VBA-on-GitHub/blob/main/gitattributes/CRLF%20everywhere/.gitattributes)

This template will make sure that Git doesn't perform any line endings or text encoding conversions.
> [!WARNING]  
> If you've already commited to your repo before starting to use this template and your previous `.gitattributes` was using `* text=auto`, you will also have to perform a second step after introducing the template: line re-normalization of your code files. Otherwise, VBA files will still have the wrong line endings inside Git. To do that, you could re-commit all your VBA code files using the command `git add . --renormalize` ([read more](https://docs.github.com/en/get-started/git-basics/configuring-git-to-handle-line-endings#refreshing-a-repository-after-changing-line-endings)) or use [Enforce-CRLF](https://github.com/DecimalTurn/Enforce-CRLF) for instance.

## Explanations

If you need a refresher or you've never had to think about how line endings work, I'd suggest having a look at the introduction of [this article](https://www.hanselman.com/blog/carriage-returns-and-line-feeds-will-ultimately-bite-you-some-git-tips) by Scott Hanselman. It will explain the origin of this [CRLF][CRLFDEF] vs [LF][LFDEF] issue.

### What does `* text=auto` actually do?

When creating a `.gitattributes`, a common practice is to include `* text=auto` at the top of your file. This can be useful for certain cross-platform projects, but it's not really useful for VBA projects and it can even be a problem if that's the only entry you have in your `.gitattributes` file.

`* text=auto` is used to let Git decide automatically (`auto`) for all files (`*`) if the `text` attribute should be "set" (aka. true) or "unset" (aka. false). Having the `text` attribute as set enables end-of-line (EOL) conversion: when a matching file is added to the Git index, the file's line endings are normalized to LF<sup>1</sup>. 

### Should you include `* text=auto` in your .gitattributes?

For VBA projects, the answer is (probably) no. In the suggested template, we don't include `* text=auto` because having this as the default for all file types requires you to be careful and override every case where this behavior can cause problems. All that without seeing much benefits.

### What are the dangers of using `* text=auto`?

Some people warn against using `* text=auto` by saying that it could corrupt binary files, but usually Git is pretty good at determining if a file is a text or binary file meaning that this shouldn't be something that worries you. The real worry comes from VBA's peculiarities.

The VBE's heyday was in the 90s. Back then, Windows was running on CRLF and there were no intentions of supporting that competing LF standard used in Unix systems. Windows has now moved on and even Notepad supports LF nowadays, but the VBE does not, sadly.

This means that the VBE expects files to use CRLF and if you try to import a VBA code file with LF, you'll experience weird bugs (especially with `.frm` and `.cls` files) such as the one described [here](https://github.com/VBA-tools/VBA-Dictionary/issues/5), [here](https://github.com/VBA-tools/VBA-Dictionary/issues/12), [here](https://github.com/VBA-tools/VBA-Dictionary/issues/38), [here](https://github.com/VBA-tools/VBA-JSON/issues/265), [here](https://github.com/cristianbuse/VBA-UserForm-MouseScroll/issues/28), [here](https://www.reddit.com/r/vba/comments/1ddpvtb/comment/l875ps5/), [here](https://github.com/VBA-tools/VBA-Web/issues/24) or [here](https://www.excelforum.com/excel-programming-vba-macros/1317233-importing-a-downloaded-from-file-from-web-is-not-working.html). For that reason, the recommended template `.gitattributes` file in this repository prevents Git from performing LF normalization on VBA code files. This means that CRLFs will be preserved since you've likely exported them via the VBE which exports them using CRLF.

All the issues mentioned in the previous paragraph have a common cause: People tried to download a single VBA code file, but GitHub gave them the "raw" file stored inside Git which included LF instead of CRLF. Indeed, the conversion back to CRLF is only taken care of when you clone or downloading the repository as a `.zip` file<sup>2</sup>. In the end, if we don't want newcomers to face any of theses issues, the simplest option is to prevent all line normalization, just like what is suggested here.

And if you are confident that you have addressed all problematic cases including the ones in the `.gitattributes` template and still want to add `* text=auto`, feel free to do so, but it's important to place that line at the top, so that the lines that come after can override this behavior<sup>3</sup>.

### What about VBA projects with macOS support?

Even if your VBA project has support for macOS, the Mac version of the VBE still expects the code files that it imports to use CRLF.

### Does the use of `-text` affect how Git performs diffs?

The short answer is no, this won't affect how Git performs diffs.

The [`text` attribute](https://git-scm.com/docs/gitattributes#_text) is only there to determine if Git will perform line endings conversion, the [`diff` attribute](https://git-scm.com/docs/gitattributes#_generating_diff_text) is the one to determine how diffs are performed. 

`-text` is not equivalent to the `binary` attribute. The `binary` attribute is a [macro attribute](https://git-scm.com/docs/gitattributes#_using_macro_attributes) equivalent to unsetting the 3 attributes:
```
-text -diff -merge
```
Setting the `binary` attribute will indeed prevent diffs from being performed since it includes `-diff`, not because of `-text`.

### Should you specify the `working-tree-encoding` attribute?

Probably no, unless you have code comments in a language other than English.

You should avoid non-ASCII characters in code that will run or will be shared outside your computer because those characters will appear differently on someone with a machine using a different encoding due to language configurations. This is especially important for text inside quotes because that will actually change the behavior of your code!

If you do have non-ASCII characters in your comments, a benefit of using `working-tree-encoding` is that people will have a better experience interacting with your code on GitHub. Indeed, GitHub has some issues when dealing with non-UTF8 encoding ([example](https://github.com/orgs/community/discussions/77064)).

<details>
<summary>How to specify the working-tree-encoding?</summary>

To specify the encoding, you'd need to specify something like this:
```
*.bas -text diff working-tree-encoding=CP1252
```

In this case, I'm using CP1252 which is the usual Windows code page for Windows OS in North American and Western Europe, but it might differ on your system. To get the number that goes after "cp" on your local machine, you can run the following Powershell command :
```
Get-WinSystemLocale | Select-Object @{ n='ANSI Code Page';   e={ $_.TextInfo.AnsiCodePage } }
```
[<sup>source</sup>](https://serverfault.com/questions/80635/how-can-i-manually-determine-the-codepage-and-locale-of-the-current-os/836221#836221)

</details>

## Should you specify the `linguist-language` attribute?

Regarding the attribute `linguist-language=vba`, I choose not to include it in the suggested template because I believe that GitHub's system to detect if a file is VBA or not is decent enough and if your files aren't detected as VBA, it might be telling you something about your code.

To mitigate misindentification, make sure that:
  - Your files include the VBE's metadata such as the `Attribute VB_Name = "..."`
  - You use the right file extensions (bas, frm or vba)

If your `.bas` files are still marked as VB6, it's likely that they don't contain any VBA specific syntax, so maybe it's OK if is classified as VB6 instead 🤷. However, if you end up with the majority of you code identified as VB6, you can have a look a this [troubleshooting guide](https://github.com/DecimalTurn/VBA-on-GitHub/blob/main/TROUBLESHOOTING.md).

# Should you include a .editorconfig file in your repository?

The `.editorconfig` file is used by many popular IDE and text editor to specify the default behavior when dealing with certain text files. If you want to make it easier for people to use other editors than the VBE, you can always include that file in your project.

# .editorconfig

A template .editorconfig file is provided here: 

📖 [Template](https://github.com/DecimalTurn/VBA-on-GitHub/blob/main/editorconfig/.editorconfig)

<hr>

**Footnotes**

[1] - [Git - gitattributes Documentation](https://git-scm.com/docs/gitattributes#_text)

[2] - Note that using `* text=auto` without specifying `-text` or at least `eol=clrf` for `.frm` and `.cls` files will mean that the `.zip` file produced by GitHub when downloading a repo won't have their line endings converted back to CRLF. See for instance [here](https://www.vbforums.com/showthread.php?910543-Github-corrupted-FRM-file-what-settings-do-you-have-in-gitattributes&p=5676869#post5676869) or [here](https://github.com/joyfullservice/msaccess-vcs-addin/issues/150).

[3] - [Mind the End of Your Line ∙ Adaptive Patchwork](https://adaptivepatchwork.com/2012/03/01/mind-the-end-of-your-line/)

[VBEDEF]: ## "Visual Basic Editor"

[CRLFDEF]: ## "Carriage Return + Line Feed"

[LFDEF]: ## "Line Feed (only)"

 
