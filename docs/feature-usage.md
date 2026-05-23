Fluent UI Components Usage
This page documents how to use the custom Fluent‑style components available in the theme: callouts, details, tabs, code blocks, terminal blocks, tables, and zoomable images.

1. Callouts

1.1 Note callout
<!-- Note callout Partial -->
{% capture body %}
This is a simple note callout.
{% endcapture %}

{% include callout-note.html
  title="Note"
  content=body
%}

1.2 Note with code block
<!-- Caution callout with code block Partial -->
{% capture body %}
You can run this macro:

Debug.Print "Hello"

text
{% endcapture %}

{% include callout-note.html
  title="Example macro"
  content=body
%}

1.3 Caution callout
<!-- Caution callout Partial -->
{% capture caution_body %}
This action modifies all open documents. Save your work first.
{% endcapture %}

{% include callout-caution.html
  title="Before you proceed"
  content=caution_body
%}

1.4 Danger callout

<!-- Danger callout Partial -->
{% capture danger_body %}
This will delete files permanently. Make sure you have a backup.
{% endcapture %}

{% include callout-danger.html
  title="Data loss warning"
  content=danger_body
%}

2. Details (expand / collapse)

<!-- Details Partial -->
{% capture details_body %}
This section contains additional explanation and examples.

Sub Example()
Debug.Print "Inside details block"
End Sub

text
{% endcapture %}

{% include details.html
  summary="Click to expand"
  content=details_body
%}
To open by default:

<!-- Details Partial, Open by Default -->
{% include details.html
  summary="Always open"
  content=details_body
  open="true"
%}

3. Tabs with code blocks
Example with three language tabs using the simplified tabs include.

<!-- Tabs Partial -->
{% capture vba_code %}
Sub HelloWorld()
Debug.Print "Hello from VBA"
End Sub

text
{% endcapture %}

{% capture csharp_code %}
public void HelloWorld()
{
Console.WriteLine("Hello from C#");
}

text
{% endcapture %}

<!-- Multiple Tabs Partial -->
{% capture cpp_code %}
void HelloWorld() {
std::cout << "Hello from C++" << std::endl;
}

{% endcapture %}

{% include tabs-simple.html
  id="hello-tabs"
  tab1="VBA"
  tab2="C#"
  tab3="C++"
  content1=vba_code
  content2=csharp_code
  content3=cpp_code
%}

4. Code blocks
4.1 Enhanced code block with filename and optional line numbers

<!-- Code Block Partial -->
{% capture code_content_3 %}

```xml
<Package Name="TestApplication.MSI" 
         Manufacturer="TODO Manufacturer" 
         Version="1.0.0.0" 
         UpgradeCode="f3873438-14af-4dcb-9ba2-4b474e155f08">

</Package>
```

{% endcapture %}

{% include code-block.html
  content=code_content_3
  filename="[Package] Element"
  lines="true"
%}

The copy button in the header copies the full code content to the clipboard.

5. Terminal blocks

<!-- Code Terminal Partial -->
{% capture terminal_output %}
npm install
bundle exec jekyll serve
Server address: http://127.0.0.1:4000
Server running... press Ctrl+C to stop.
{% endcapture %}

{% include code-terminal.html
  id="install-terminal"
  title="Terminal"
  content=terminal_output
%}
The terminal header includes a copy button that copies all terminal lines.

6. Tables (Fluent style)

<!-- Fluent Table Partial -->
{% include fluent-table.html
  caption="Included Files"
  headers="File Name|Purpose"
  rows="
`OpenAssociatedDrawing.bas`|The macro module to import
`README.md`|User guide in Markdown format
"
%}
For a compact table:

<!-- Compact Fluent Table Partial -->
{% include fluent-table.html
  caption="Macro difficulty levels"
  headers="Level|Description"
  rows="
Beginner|Simple recording or small changes
Intermediate|API usage + error handling
Advanced|Automation across many documents
"
  compact="true"
%}

7. Zoomable images (lightbox)

<!-- Zoomable images Partial -->
{% include image-zoom.html
  src="/assets/images/export-dialog.png"
  alt="Export to PDF dialog"
  caption="Figure 1: Export to PDF options in SOLIDWORKS"
%}
Clicking the image opens a Fluent‑style lightbox with the same image and caption.

<!-- Link Partial -->
{% include link-new-tab.html href="/download-inventor-macros/" text="**⬇ Download Now — It’s 100% Free**" %}

<!-- Video Partial -->
{% include video-block.html 
  video_id="9q4BhjBUgYE"
  caption="Export all open Autodesk Inventor drawings to PDF automatically with a single macro."
%}