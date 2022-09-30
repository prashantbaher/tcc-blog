"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[3463],{3905:(e,t,r)=>{r.d(t,{Zo:()=>p,kt:()=>h});var o=r(67294);function n(e,t,r){return t in e?Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}):e[t]=r,e}function a(e,t){var r=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),r.push.apply(r,o)}return r}function l(e){for(var t=1;t<arguments.length;t++){var r=null!=arguments[t]?arguments[t]:{};t%2?a(Object(r),!0).forEach((function(t){n(e,t,r[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(r)):a(Object(r)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(r,t))}))}return e}function i(e,t){if(null==e)return{};var r,o,n=function(e,t){if(null==e)return{};var r,o,n={},a=Object.keys(e);for(o=0;o<a.length;o++)r=a[o],t.indexOf(r)>=0||(n[r]=e[r]);return n}(e,t);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);for(o=0;o<a.length;o++)r=a[o],t.indexOf(r)>=0||Object.prototype.propertyIsEnumerable.call(e,r)&&(n[r]=e[r])}return n}var s=o.createContext({}),m=function(e){var t=o.useContext(s),r=t;return e&&(r="function"==typeof e?e(t):l(l({},t),e)),r},p=function(e){var t=m(e.components);return o.createElement(s.Provider,{value:t},e.children)},u={inlineCode:"code",wrapper:function(e){var t=e.children;return o.createElement(o.Fragment,{},t)}},d=o.forwardRef((function(e,t){var r=e.components,n=e.mdxType,a=e.originalType,s=e.parentName,p=i(e,["components","mdxType","originalType","parentName"]),d=m(r),h=n,c=d["".concat(s,".").concat(h)]||d[h]||u[h]||a;return r?o.createElement(c,l(l({ref:t},p),{},{components:r})):o.createElement(c,l({ref:t},p))}));function h(e,t){var r=arguments,n=t&&t.mdxType;if("string"==typeof e||n){var a=r.length,l=new Array(a);l[0]=d;var i={};for(var s in t)hasOwnProperty.call(t,s)&&(i[s]=t[s]);i.originalType=e,i.mdxType="string"==typeof e?e:n,l[1]=i;for(var m=2;m<a;m++)l[m]=r[m];return o.createElement.apply(null,l)}return o.createElement.apply(null,r)}d.displayName="MDXCreateElement"},3470:(e,t,r)=>{r.r(t),r.d(t,{assets:()=>s,contentTitle:()=>l,default:()=>u,frontMatter:()=>a,metadata:()=>i,toc:()=>m});var o=r(87462),n=(r(67294),r(3905));const a={title:"VBA UserForms",tags:["VBA"],permalink:"/vba/userform/"},l=void 0,i={unversionedId:"vba-userform",id:"vba-userform",title:"VBA UserForms",description:"A UserForm is useful if your VBA macro needs to get information from a user.",source:"@site/docs/vba/27-vba-userform.md",sourceDirName:".",slug:"/vba-userform",permalink:"/vba/vba-userform",draft:!1,tags:[{label:"VBA",permalink:"/vba/tags/vba"}],version:"current",sidebarPosition:27,frontMatter:{title:"VBA UserForms",tags:["VBA"],permalink:"/vba/userform/"},sidebar:"tutorialSidebar",previous:{title:"VBA GetOpenFilename, GetSaveAsFilename and Getting a Folder Name",permalink:"/vba/vba-other-dialog"},next:{title:"VBA Userforms - Open new Part document",permalink:"/vba/open-part-from-userform"}},s={},m=[{value:"Userforms Working",id:"userforms-working",level:2},{value:"Inserting a new UserForm",id:"inserting-a-new-userform",level:2},{value:"Adding controls to a UserForm",id:"adding-controls-to-a-userform",level:2},{value:"Changing properties for a UserForm control",id:"changing-properties-for-a-userform-control",level:2},{value:"Viewing the UserForm Code window",id:"viewing-the-userform-code-window",level:2},{value:"Showing the UserForm",id:"showing-the-userform",level:2},{value:"Using information from a UserForm",id:"using-information-from-a-userform",level:2},{value:"Thank you!!!!",id:"thank-you",level:3},{value:"UPDATE:",id:"update",level:2}],p={toc:m};function u(e){let{components:t,...a}=e;return(0,n.kt)("wrapper",(0,o.Z)({},p,a,{components:t,mdxType:"MDXLayout"}),(0,n.kt)("p",null,"A ",(0,n.kt)("em",{parentName:"p"},"UserForm")," is useful if your VBA macro needs to get information from a user. "),(0,n.kt)("p",null,"For example, your macro may have some options that can be specified in a UserForm. "),(0,n.kt)("p",null,"If only a few pieces of information are required (for example, a ",(0,n.kt)("em",{parentName:"p"},"Yes/No")," answer or a text ",(0,n.kt)("em",{parentName:"p"},"string"),"), one of the techniques I describe in previous articles may do the job. "),(0,n.kt)("p",null,"But if you need to obtain more information, you must create a UserForm."),(0,n.kt)("p",null,"To create a UserForm, you usually take the following general steps:"),(0,n.kt)("ul",null,(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"Determine how the dialog box will be used and where it will be displayed in your VBA macro.")),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"Activate the VBE and insert a new UserForm object. A UserForm object holds a single UserForm.")),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"Add controls to the UserForm. Controls include items such as text boxes, buttons, check boxes, and list boxes.")),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"Use the Properties window to modify the properties for the controls or for the UserForm itself.")),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"Write ",(0,n.kt)("em",{parentName:"p"},"event-handler")," procedures for the controls (for example, a macro that executes when the user clicks a button in the dialog box). These procedures are stored in the Code window for the UserForm object.")),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"Write a ",(0,n.kt)("em",{parentName:"p"},"procedure")," (stored in a VBA module) that displays the dialog box to the user."))),(0,n.kt)("p",null,"When you are designing a ",(0,n.kt)("em",{parentName:"p"},"UserForm"),", you are creating what developers call the ",(0,n.kt)("strong",{parentName:"p"},"Graphical User Interface (GUI)")," to your application. "),(0,n.kt)("p",null,"Take some time to consider what your form should look like and how your users are likely to want to interact with the elements on the UserForm. "),(0,n.kt)("p",null,"Try to guide them through the steps they need to take on the form by carefully considering the arrangement and wording of the controls. "),(0,n.kt)("p",null,"Like most things VBA-related, the more you do it, the easier it gets."),(0,n.kt)("h2",{id:"userforms-working"},"Userforms Working"),(0,n.kt)("p",null,"Each dialog box that you create is stored in its own UserForm object \u2014 one dialog box per UserForm. "),(0,n.kt)("p",null,"You create and access these UserForms in the Visual Basic Editor."),(0,n.kt)("h2",{id:"inserting-a-new-userform"},"Inserting a new UserForm"),(0,n.kt)("p",null,"To insert a UserForm object with the following steps:"),(0,n.kt)("ol",null,(0,n.kt)("li",{parentName:"ol"},"In the macro, you can insert User form with following 2 ways:")),(0,n.kt)("ul",null,(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},'From "Menu Bar" -> "UserForm"')),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"From \u201cStandard Toolbar\u201d by clicking \u201cInsert UserForm\u201d ",(0,n.kt)("img",{alt:"A-new-userform-object",src:r(30470).Z,width:"37",height:"28"})),(0,n.kt)("p",{parentName:"li"},"The VBE insert a new UserForm object with an empty dialog box."))),(0,n.kt)("ol",{start:2},(0,n.kt)("li",{parentName:"ol"},"If \u201cProperty window\u201d is not available in your macro, press ",(0,n.kt)("inlineCode",{parentName:"li"},"F4")," to display \u201cProperty window\u201d.")),(0,n.kt)("p",null,"The VBE inserts a new UserForm object, which contains an empty dialog box."),(0,n.kt)("p",null,"Below figure shows a UserForm \u2014 an empty dialog box with some controls in Toolbox."),(0,n.kt)("p",null,(0,n.kt)("img",{alt:"Empty-userform-object",src:r(84168).Z,width:"1366",height:"741"})),(0,n.kt)("h2",{id:"adding-controls-to-a-userform"},"Adding controls to a UserForm"),(0,n.kt)("p",null,"When you activate a UserForm, the VBE displays the Toolbox in a floating window, as shown in above figure. "),(0,n.kt)("p",null,"You use the tools in the Toolbox to add controls to your UserForm. "),(0,n.kt)("p",null,"If the Toolbox doesn\u2019t appear when you activate your UserForm, choose ",(0,n.kt)("strong",{parentName:"p"},"View -> Toolbox"),"."),(0,n.kt)("p",null,"To add a control, just click the desired control in the Toolbox and drag it into the dialog box to create the control. "),(0,n.kt)("p",null,"After you add a control, you can move and resize it by using standard techniques."),(0,n.kt)("p",null,"Below table indicates the various tools, as well as their capabilities. "),(0,n.kt)("p",null,"To determine which tool is which, hover your mouse pointer over the control and read the small pop-up description."),(0,n.kt)("table",null,(0,n.kt)("thead",{parentName:"table"},(0,n.kt)("tr",{parentName:"thead"},(0,n.kt)("th",{parentName:"tr",align:null},"Controls"),(0,n.kt)("th",{parentName:"tr",align:null},"What it does"))),(0,n.kt)("tbody",{parentName:"table"},(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"Label"),(0,n.kt)("td",{parentName:"tr",align:null},"Shows text")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"TextBox"),(0,n.kt)("td",{parentName:"tr",align:null},"Determines which of the file filters the dialog box displays by default.")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"ComboBox"),(0,n.kt)("td",{parentName:"tr",align:null},"Display a drop-down list.")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"ListBox"),(0,n.kt)("td",{parentName:"tr",align:null},"Display a list of items.")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"CheckBox"),(0,n.kt)("td",{parentName:"tr",align:null},"Useful for On/off or Yes/No options.")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"OptionButton"),(0,n.kt)("td",{parentName:"tr",align:null},"Used in groups; allows the user to select one of several options.")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"ToggleButoon"),(0,n.kt)("td",{parentName:"tr",align:null},"A button that is either on or off.")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"Frame"),(0,n.kt)("td",{parentName:"tr",align:null},"A container for other control.")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"CommandButton"),(0,n.kt)("td",{parentName:"tr",align:null},"A clickable button.")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"TabStrip"),(0,n.kt)("td",{parentName:"tr",align:null},"Display Tabs")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"MultiPage"),(0,n.kt)("td",{parentName:"tr",align:null},"A tabbed container for other objects.")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"ScrollBar"),(0,n.kt)("td",{parentName:"tr",align:null},"A draggable bar.")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"SpinButton"),(0,n.kt)("td",{parentName:"tr",align:null},"A clickable button often used for changing a value.")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"Image"),(0,n.kt)("td",{parentName:"tr",align:null},"Contains an image")),(0,n.kt)("tr",{parentName:"tbody"},(0,n.kt)("td",{parentName:"tr",align:null},"RefEdit"),(0,n.kt)("td",{parentName:"tr",align:null},"Allows the user to select a range.")))),(0,n.kt)("h2",{id:"changing-properties-for-a-userform-control"},"Changing properties for a UserForm control"),(0,n.kt)("p",null,"Every control you add to a UserForm has a number of properties that determine how the control looks or behaves. "),(0,n.kt)("p",null,"In addition, the UserForm itsel also has its own set of properties. "),(0,n.kt)("p",null,"You can change these properties with the ",(0,n.kt)("em",{parentName:"p"},"Properties window"),". "),(0,n.kt)("p",null,"Below figure shows the properties window when a ",(0,n.kt)("inlineCode",{parentName:"p"},"CommandButton")," control is selected:"),(0,n.kt)("p",null,(0,n.kt)("img",{alt:"Empty-userform-object",src:r(61455).Z,width:"1366",height:"739"})),(0,n.kt)("p",null,"Properties for controls include the following:"),(0,n.kt)("ul",null,(0,n.kt)("li",{parentName:"ul"},"Name"),(0,n.kt)("li",{parentName:"ul"},"Width"),(0,n.kt)("li",{parentName:"ul"},"Height"),(0,n.kt)("li",{parentName:"ul"},"Value"),(0,n.kt)("li",{parentName:"ul"},"Caption")),(0,n.kt)("p",null,"Each control has its own set of properties (although many controls have some common properties). To change a property using the Properties window:"),(0,n.kt)("ol",null,(0,n.kt)("li",{parentName:"ol"},"Make sure that the correct control is selected in the UserForm."),(0,n.kt)("li",{parentName:"ol"},"Make sure the Properties window is visible (press ",(0,n.kt)("inlineCode",{parentName:"li"},"F4")," if it\u2019s not)."),(0,n.kt)("li",{parentName:"ol"},"In the Properties window, click on the property that you want to change."),(0,n.kt)("li",{parentName:"ol"},"Make the change in the right portion of the Properties window.")),(0,n.kt)("p",null,"If you select the ",(0,n.kt)("strong",{parentName:"p"},"UserForm")," itself (not a ",(0,n.kt)("strong",{parentName:"p"},"control")," on the UserForm), you can use the Properties window to adjust UserForm properties"),(0,n.kt)("blockquote",null,(0,n.kt)("p",{parentName:"blockquote"},"Some of the UserForm properties serve as default settings for new controls you drag onto the UserForm. For example, if you change the Font property for a UserForm, controls that you add will use that same font. Controls that are already on the UserForm are not affected.")),(0,n.kt)("h2",{id:"viewing-the-userform-code-window"},"Viewing the UserForm Code window"),(0,n.kt)("p",null,"Every UserForm object has a Code module that holds the VBA code (",(0,n.kt)("em",{parentName:"p"},"the event-handler procedures"),") executed when the user works with the dialog box. "),(0,n.kt)("p",null,"To view the Code module, press ",(0,n.kt)("inlineCode",{parentName:"p"},"F7"),". "),(0,n.kt)("p",null,"The ",(0,n.kt)("em",{parentName:"p"},"Code window")," is empty until you add some procedures. Press ",(0,n.kt)("inlineCode",{parentName:"p"},"Shift+F7")," to return to the dialog box."),(0,n.kt)("p",null,"Here\u2019s another way to switch between the Code window and the UserForm display: "),(0,n.kt)("ul",null,(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"Use the View Code and View Object buttons in the Project window\u2019s title bar. ")),(0,n.kt)("li",{parentName:"ul"},(0,n.kt)("p",{parentName:"li"},"Or right-click the UserForm and choose View Code. "))),(0,n.kt)("p",null,"If you\u2019re viewing code, ",(0,n.kt)("em",{parentName:"p"},"double-click")," the UserForm name in the Project window to return to the UserForm."),(0,n.kt)("h2",{id:"showing-the-userform"},"Showing the UserForm"),(0,n.kt)("p",null,"You display a UserForm by using the UserForm\u2019s ",(0,n.kt)("inlineCode",{parentName:"p"},"Show")," method in a VBA procedure."),(0,n.kt)("p",null,"The macro that displays the dialog box must be in a VBA module \u2014 not in the Code window for the UserForm."),(0,n.kt)("p",null,"The following procedure displays the dialog box named ",(0,n.kt)("inlineCode",{parentName:"p"},"UserForm1"),":"),(0,n.kt)("pre",null,(0,n.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Showing the UserForm\nSub ShowDialogBox()\n  UserForm.Show\n  'Other statements can go here\nEnd Sub\n")),(0,n.kt)("p",null,"When Solidworks displays the dialog box, the ",(0,n.kt)("inlineCode",{parentName:"p"},"ShowDialogBox")," macro halts until the user closes the dialog box. "),(0,n.kt)("p",null,"Then VBA executes any remaining statements in the procedure. "),(0,n.kt)("p",null,"Most of the time, you won\u2019t have any more code in the procedure."),(0,n.kt)("h2",{id:"using-information-from-a-userform"},"Using information from a UserForm"),(0,n.kt)("p",null,"The VBE provides a name for each control you add to a UserForm. "),(0,n.kt)("p",null,"The control\u2019s name corresponds to its ",(0,n.kt)("inlineCode",{parentName:"p"},"Name")," property. "),(0,n.kt)("p",null,"Use this name to refer to a particular control in your code. "),(0,n.kt)("p",null,"For example, if you add a ",(0,n.kt)("inlineCode",{parentName:"p"},"CheckBox")," control to a UserForm named ",(0,n.kt)("inlineCode",{parentName:"p"},"UserForm1"),", the CheckBox control is named ",(0,n.kt)("inlineCode",{parentName:"p"},"CheckBox1")," by default. "),(0,n.kt)("p",null,"The following statement makes this control appear with a check mark:"),(0,n.kt)("pre",null,(0,n.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"UserForm1.CheckBox1.Value = True\n")),(0,n.kt)("p",null,"Most of the time, you write the code for a UserForm in the UserForm\u2019s code module. "),(0,n.kt)("p",null,"If that\u2019s the case, you can omit the UserForm object qualifier and write the statement like this:"),(0,n.kt)("pre",null,(0,n.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"CheckBox1.Value = True\n")),(0,n.kt)("blockquote",null,(0,n.kt)("p",{parentName:"blockquote"},"I recommend that you change the default name the VBE has given to your controls to something more meaningful.")),(0,n.kt)("p",null,"This will sum-up our tutorials on Visual Basic for Application. From now on I will give tutorials on how to use Solidworks commands with the help of VBA Macro."),(0,n.kt)("p",null,"If you want to know any explaination on any topic related to VBA, please drop a comment and I will try to give it to you. "),(0,n.kt)("h3",{id:"thank-you"},"Thank you!!!!"),(0,n.kt)("h2",{id:"update"},"UPDATE:"),(0,n.kt)("p",null,"I have started VBA UserForm Example in this tutorials lists. "),(0,n.kt)("p",null,"So if you want to learn how I use these Forms, you can watch them in UserForm Example List Post."))}u.isMDXComponent=!0},84168:(e,t,r)=>{r.d(t,{Z:()=>o});const o=r.p+"assets/images/1.Anewuserformobject-c638eb84cfb4e3b46b2485935168a618.PNG"},61455:(e,t,r)=>{r.d(t,{Z:()=>o});const o=r.p+"assets/images/2.UsethePropertiesWindowstoChangethePropertiesofUserFormControls-e41f74f25a7b3903776ddaa82d5a8028.PNG"},30470:(e,t,r)=>{r.d(t,{Z:()=>o});const o="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACUAAAAcCAYAAADm63ZmAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAJoSURBVFhH7ZerjttAFIbnEfIIeQTDhYaGhoEDAwcaDgwcVAUODDRYsKyWSkIqWUWVlrisBZWc+w2c/mdsJ9FmnHVbddeglj7NxI7OfPnnJI5FWZb01iwWi7uI0+lEfeO/VFfE8XikvvGqVGFS+vLhkfJJCizlGiSWMmUoGwOpKRtpSuOaSDl8tbrSKqVt4UaW+vb4iYo0q5gBm1E+hSSuORrZWtgE8qbe7yAOhwP50NPCiblFI6RRkzqQRlhhmQd5xjCQ8tXsSquUghQCIWUKKvfkP3Cer50pK4LoX0lBxs6JzBNRPKnFGBwvRYorsrz4e6n9fk8+JEQaKZ1CLMlJodllYpyMNpl7rdBHfE4qTcX3i5SvZlfapXQlFSrIzLCNthJrEtLca18LKhBPgbjk+A2k4qTqqUBmFEMsnhDJ6UWsSU3ip8EBqZyl5velhBA33Lxnt9uRD16ct24YISIcnFiskYgpK7GmjyDCsNC1lK9mw7WQ97rvJMMS2CEaPBgahiBOXWohEmQx/iJkeUnp/AXZ61JMmxAjttst+WAp9LBLJ0yAwmI1PJeTKjEW8OGr2RWx2WzIB0udGWNxpBSMUiRm3ZYOkJ4DSf74ufTW+FNapZjcDh2puYwZjxOM0yFZPcB8QDYBmBsFMDdKeOt1RazXa2qDRXBDczKXEbcXSOFO7EToKXZSNAudVGkDkrHw1uuKWK1W1AanwjIXFNEcQMgBIfw1cEIs45gGFIfCW68r96WwRfzp9bhGClLMqIITcUQVLNPgq9cVsVwuqW8I39PEe9NPKd9z2XvTz576+PmZ+kYPpZ7pF0ZHakU/W9NSAAAAAElFTkSuQmCC"}}]);