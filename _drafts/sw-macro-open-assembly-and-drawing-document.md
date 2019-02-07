# Open Assembly and Drawing document with Solidworks VBA Macro

In this post, we see how to open following documents with *Solidworks VBA macro*:

1. *Solidworks Assembly document*

2. *Solidworks Drawing document*

    * **Without** Pre-defined Sheet size

    * **With** Pre-defined Sheet size

    * *With Custom Sheet size*

## Open Solidworks Assembly Document

The code for opening *default Assembly document* is identical to the *default Part template* with only one change.

First, let us see the code to open default Assembly document.

```vb
Option Explicit

' Creating variable for Solidworks application
Dim swApp As SldWorks.SldWorks
' Creating variable for Solidworks document
Dim swDoc As SldWorks.ModelDoc2

' Main function of our VBA program
Sub main()

    ' Setting Solidworks variable to Solidworks application
    Set swApp = Application.SldWorks
    
    ' Creating string type variable for storing default Assembly location
    Dim defaultTemplate As String
    ' Setting value of this string type variable to "Default Assembly template"
    defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateAssembly)

    ' Setting Solidworks document to new Assembly document
    Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)

End Sub
```

As you can see in the above code and the code is given in the previous post is almost identically.

In case you have not read my previous post (Solidworks Macros - Open new Part document), I recommend you to read that post first. 

I have already explained each and every line in this code there.


