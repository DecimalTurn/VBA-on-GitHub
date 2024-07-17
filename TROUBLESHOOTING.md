## What to do when VBA code in `.bas` files is detected as VB6 on GitHub?

Sometimes, VBA code can be detected as a Visual Basic 6.0 (VB6).

![image](https://github.com/ClimberMel/SMF_Add-in/assets/31558169/36fea138-cc22-405f-8db8-077517d2e559)

The reason for it is that a some code modules in the project don't have any VBA specific syntax and [Github-Linguist](https://github.com/github-linguist/linguist) (the GitHub service that calculates the language statistics) looks only at the content and extension of a file to determine the language used for that file.

To solve this issue and make sure that files in this repo are correctly identified as VBA, there are a mainly 3 options:

### A. Add the language information as a comment at the top of the file. 
(This is the best option in my opinion.)
ie.:

```vba
Attribute VB_Name = "ComponentName"
'@Lang VBA
```

You can automate this process by using the [VBA-Language-Annotation](https://github.com/DecimalTurn/VBA-Language-Annotation) action.

### B. Add the following in the `.gitattributes` file:

```gitattributes

*.bas linguist-language=vba

```

The advantage of this method is that you only have to make the change in one place. The disadvantage is that this won't solve the problem for GitHub searches because the `linguist-language` attribute is not supported for searches [^1]. This means that if the VBA language filter is used in a search query, files that are currently labelled as VB6 still won't show up.

### C. Change the file extension from `.bas` to `.vba`

That second extension avoids any ambiguity between VBA and VB6, but you won't be able to use the simple Export feature of the VBE since you'll need to perform the renaming on every export.

[^1]: This is due to the fact that [Github Search relies on Go-Enry](https://github.com/github-linguist/linguist/blob/master/docs/troubleshooting.md#:~:text=GitHub%20Search%20relies%20on%20go%2Denry%20which%20doesn%27t%20support%20overrides%20at%20the%20moment%20(More%20info).) and [Go-Enrs doesn't support overrides at the moment.](https://github.com/go-enry/go-enry#:~:text=Overriding%20languages%20and%20types%20though%20.gitattributes%20is%20not%20yet%20supported.%20See%20%2318.)
