"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[3707],{3905:function(e,t,n){n.d(t,{Zo:function(){return p},kt:function(){return m}});var o=n(67294);function a(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function i(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var o=Object.getOwnPropertySymbols(e);t&&(o=o.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,o)}return n}function l(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?i(Object(n),!0).forEach((function(t){a(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):i(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function r(e,t){if(null==e)return{};var n,o,a=function(e,t){if(null==e)return{};var n,o,a={},i=Object.keys(e);for(o=0;o<i.length;o++)n=i[o],t.indexOf(n)>=0||(a[n]=e[n]);return a}(e,t);if(Object.getOwnPropertySymbols){var i=Object.getOwnPropertySymbols(e);for(o=0;o<i.length;o++)n=i[o],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(a[n]=e[n])}return a}var s=o.createContext({}),d=function(e){var t=o.useContext(s),n=t;return e&&(n="function"==typeof e?e(t):l(l({},t),e)),n},p=function(e){var t=d(e.components);return o.createElement(s.Provider,{value:t},e.children)},c={inlineCode:"code",wrapper:function(e){var t=e.children;return o.createElement(o.Fragment,{},t)}},u=o.forwardRef((function(e,t){var n=e.components,a=e.mdxType,i=e.originalType,s=e.parentName,p=r(e,["components","mdxType","originalType","parentName"]),u=d(n),m=a,k=u["".concat(s,".").concat(m)]||u[m]||c[m]||i;return n?o.createElement(k,l(l({ref:t},p),{},{components:n})):o.createElement(k,l({ref:t},p))}));function m(e,t){var n=arguments,a=t&&t.mdxType;if("string"==typeof e||a){var i=n.length,l=new Array(i);l[0]=u;var r={};for(var s in t)hasOwnProperty.call(t,s)&&(r[s]=t[s]);r.originalType=e,r.mdxType="string"==typeof e?e:a,l[1]=r;for(var d=2;d<i;d++)l[d]=n[d];return o.createElement.apply(null,l)}return o.createElement.apply(null,n)}u.displayName="MDXCreateElement"},54358:function(e,t,n){n.r(t),n.d(t,{assets:function(){return p},contentTitle:function(){return s},default:function(){return m},frontMatter:function(){return r},metadata:function(){return d},toc:function(){return c}});var o=n(87462),a=n(63366),i=(n(67294),n(3905)),l=["components"],r={title:"SOLIDWORKS C# API - Open SOLIDWORKS Document",tags:["SOLIDWORKS C# API"],categories:"SOLIDWORKS-C#-API",permalink:"/solidworks-csharp/open-solidworks-document/",id:"open-solidworks-document"},s=void 0,d={unversionedId:"open-solidworks-document",id:"open-solidworks-document",title:"SOLIDWORKS C# API - Open SOLIDWORKS Document",description:"OBJECTIVE",source:"@site/docs/solidworks-csharp/001.4-open-solidworks-document.md",sourceDirName:".",slug:"/open-solidworks-document",permalink:"/solidworks-csharp/open-solidworks-document",draft:!1,tags:[{label:"SOLIDWORKS C# API",permalink:"/solidworks-csharp/tags/solidworks-c-api"}],version:"current",frontMatter:{title:"SOLIDWORKS C# API - Open SOLIDWORKS Document",tags:["SOLIDWORKS C# API"],categories:"SOLIDWORKS-C#-API",permalink:"/solidworks-csharp/open-solidworks-document/",id:"open-solidworks-document"},sidebar:"tutorialSidebar",previous:{title:"SOLIDWORKS C# API - Open SOLIDWORKS",permalink:"/solidworks-csharp/open-solidworks"},next:{title:"SOLIDWORKS C# API - Update Dimension State",permalink:"/solidworks-csharp/dimension-state"}},p={},c=[{value:"OBJECTIVE",id:"objective",level:2},{value:"DEMO VIDEO",id:"demo-video",level:2},{value:"CREATE A NEW PRISM PROJECT",id:"create-a-new-prism-project",level:2},{value:"BUILD SOLUTION",id:"build-solution",level:2},{value:"WHY WE BUILD SOLUTION ?",id:"why-we-build-solution-",level:3},{value:"ADD USER INTERFACE CONTROLS",id:"add-user-interface-controls",level:2},{value:"UPDATE WINDOW START-UP LOCATION AND HEIGHT/WIDTH",id:"update-window-start-up-location-and-heightwidth",level:3},{value:"REMOVE CONTENT CONTROL",id:"remove-content-control",level:3},{value:"ADD TEXTBLOCK FOR SELECTING DOCUMENT",id:"add-textblock-for-selecting-document",level:3},{value:"ADD COMBOBOX FOR DOCUMENTS LIST",id:"add-combobox-for-documents-list",level:3},{value:"ADD BUTTON FOR SELECTED DOCUMENT",id:"add-button-for-selected-document",level:3},{value:"UPDATE VIEWMODEL",id:"update-viewmodel",level:2},{value:"ADD DOCUMENTS LIST",id:"add-documents-list",level:3},{value:"BINDING DOCUMENT LIST TO COMBOBOX",id:"binding-document-list-to-combobox",level:3},{value:"ADD SELECTED VALUE IN COMBOBOX",id:"add-selected-value-in-combobox",level:3},{value:"ADD COMMAND TO VIEWMODEL",id:"add-command-to-viewmodel",level:3},{value:"ADD SOLIDWORKS REFERENCES",id:"add-solidworks-references",level:2},{value:"OPEN SOLIDWORKS DOCUMENT",id:"open-solidworks-document",level:2},{value:"FINAL RESULT",id:"final-result",level:2}],u={toc:c};function m(e){var t=e.components,r=(0,a.Z)(e,l);return(0,i.kt)("wrapper",(0,o.Z)({},u,r,{components:t,mdxType:"MDXLayout"}),(0,i.kt)("h2",{id:"objective"},"OBJECTIVE"),(0,i.kt)("p",null,"How to Open SOLIDWORKS Document using ",(0,i.kt)("strong",{parentName:"p"},"SOLIDWORKS C# API")," from ",(0,i.kt)("strong",{parentName:"p"},"WPF Prism Desktop Application"),"."),(0,i.kt)("p",null,"I hope you have installed ",(0,i.kt)("em",{parentName:"p"},"Visual Studio Community Edition")," on your machine."),(0,i.kt)("p",null,"If not then please go to ",(0,i.kt)("strong",{parentName:"p"},(0,i.kt)("a",{parentName:"strong",href:"/solidworks-csharp/csharp-prerequisite"},"SOLIDWORKS C# API - Prerequisite"))," post and watch the suggested videos before proceeding further."),(0,i.kt)("hr",null),(0,i.kt)("h2",{id:"demo-video"},"DEMO VIDEO"),(0,i.kt)("p",null,'Please see below video on how to "Open SOLIDWORKS Document" using ',(0,i.kt)("strong",{parentName:"p"},"SOLIDWORKS C# API")," from ",(0,i.kt)("strong",{parentName:"p"},"WPF Prism Desktop Application"),"."),(0,i.kt)("iframe",{src:"https://www.youtube.com/embed/eSgcmdkB4-8",frameborder:"0",allowfullscreen:!0,width:"100%",height:"500"}),(0,i.kt)("p",null,"Please note that there are ",(0,i.kt)("strong",{parentName:"p"},"no explanation")," in the video. "),(0,i.kt)("p",null,(0,i.kt)("strong",{parentName:"p"},"Explanation")," of each step and why we write code this way is given in this post."),(0,i.kt)("hr",null),(0,i.kt)("h2",{id:"create-a-new-prism-project"},"CREATE A NEW PRISM PROJECT"),(0,i.kt)("p",null,"In the below image I have shown you how to create a new Prism project."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(1590).Z},(0,i.kt)("img",{alt:"create-open-solidworks-document-project",src:n(35334).Z,width:"1364",height:"726"}))),(0,i.kt)("p",null,"All the steps has been already explained in ",(0,i.kt)("strong",{parentName:"p"},(0,i.kt)("a",{parentName:"strong",href:"/solidworks-csharp/open-solidworks/#create-a-new-prism-project"},"SOLIDWORKS C# API - Open SOLIDWORKS"))," article."),(0,i.kt)("p",null,"This will open a new window as shown in below image."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(44844).Z},(0,i.kt)("img",{alt:"create-new-project",src:n(37514).Z,width:"1248",height:"664"}))),(0,i.kt)("h2",{id:"build-solution"},"BUILD SOLUTION"),(0,i.kt)("p",null,'After we create our "',(0,i.kt)("em",{parentName:"p"},"OpenSolidworksDocument"),'" project, we need to select "Build Solution" option as shown in below image.'),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(64865).Z},(0,i.kt)("img",{alt:"build-solution",src:n(76216).Z,width:"976",height:"665"}))),(0,i.kt)("h3",{id:"why-we-build-solution-"},"WHY WE BUILD SOLUTION ?"),(0,i.kt)("p",null,"We build our solution because we want to make sure everything is working and there are no broken references."),(0,i.kt)("p",null,"Below image show ",(0,i.kt)("inlineCode",{parentName:"p"},"MainWindow.xaml")," file before ",(0,i.kt)("em",{parentName:"p"},"building solution"),"."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(88936).Z},(0,i.kt)("img",{alt:"before-build-solution",src:n(53095).Z,width:"1234",height:"657"}))),(0,i.kt)("p",null,"Below image show ",(0,i.kt)("inlineCode",{parentName:"p"},"MainWindow.xaml")," file after ",(0,i.kt)("em",{parentName:"p"},"building solution"),"."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(34255).Z},(0,i.kt)("img",{alt:"after-build-solution",src:n(59496).Z,width:"1234",height:"657"}))),(0,i.kt)("h2",{id:"add-user-interface-controls"},"ADD USER INTERFACE CONTROLS"),(0,i.kt)("p",null,"Below we add some UI control for user interaction."),(0,i.kt)("h3",{id:"update-window-start-up-location-and-heightwidth"},"UPDATE WINDOW START-UP LOCATION AND HEIGHT/WIDTH"),(0,i.kt)("p",null,"In below image, we update '",(0,i.kt)("em",{parentName:"p"},"Start-up location"),"' and '",(0,i.kt)("em",{parentName:"p"},"Width & Height"),"' of our window."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(49851).Z},(0,i.kt)("img",{alt:"update-window-startup-location-width-height",src:n(74608).Z,width:"1364",height:"726"}))),(0,i.kt)("p",null,"First we update ",(0,i.kt)("em",{parentName:"p"},"Start-up location of window")," by adding following line."),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-xml"},'WindowStartupLocation="CenterScreen"\n')),(0,i.kt)("p",null,"After this, we update window's ",(0,i.kt)("em",{parentName:"p"},"Height and Width")," to following values."),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-xml"},'Height="250" Width="500"\n')),(0,i.kt)("h3",{id:"remove-content-control"},"REMOVE CONTENT CONTROL"),(0,i.kt)("p",null,"In below image we remove ContentControl Tag in Grid."),(0,i.kt)("p",null,"Also, we change ",(0,i.kt)("inlineCode",{parentName:"p"},"Grid")," to ",(0,i.kt)("inlineCode",{parentName:"p"},"StackPanel")," for helding our UI Controls."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(56089).Z},(0,i.kt)("img",{alt:"remove-content-control",src:n(16701).Z,width:"1364",height:"726"}))),(0,i.kt)("h3",{id:"add-textblock-for-selecting-document"},"ADD TEXTBLOCK FOR SELECTING DOCUMENT"),(0,i.kt)("p",null,"In below image we add ",(0,i.kt)("inlineCode",{parentName:"p"},"TextBlock")," inside ",(0,i.kt)("inlineCode",{parentName:"p"},"StackPanel")," for indicating user to select a document."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(30166).Z},(0,i.kt)("img",{alt:"add-textblock",src:n(66862).Z,width:"1364",height:"726"}))),(0,i.kt)("p",null,(0,i.kt)("inlineCode",{parentName:"p"},"TextBlock")," with ",(0,i.kt)("inlineCode",{parentName:"p"},"Text")," and ",(0,i.kt)("em",{parentName:"p"},"other properties")," are given below."),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-xml"},'<TextBlock Text="Select Document"\n           Width="350"\n           Height="30"\n           Margin="25 10"\n           FontSize="25"\n           FontWeight="Medium"\n           VerticalAlignment="Center" />\n')),(0,i.kt)("h3",{id:"add-combobox-for-documents-list"},"ADD COMBOBOX FOR DOCUMENTS LIST"),(0,i.kt)("p",null,"In below image we add ",(0,i.kt)("inlineCode",{parentName:"p"},"ComboBox")," inside ",(0,i.kt)("inlineCode",{parentName:"p"},"StackPanel")," for holding our list of ",(0,i.kt)("strong",{parentName:"p"},"SOLIDWORKS documents"),"."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(4527).Z},(0,i.kt)("img",{alt:"add-documents-combobox",src:n(56342).Z,width:"1364",height:"726"}))),(0,i.kt)("p",null,(0,i.kt)("inlineCode",{parentName:"p"},"ComboBox")," with set ",(0,i.kt)("em",{parentName:"p"},"properties")," are given below."),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-xml"},'<ComboBox Width="350"\n          Height="30"\n          Margin="10"\n          VerticalAlignment="Center"\n          FontSize="16" />\n')),(0,i.kt)("h3",{id:"add-button-for-selected-document"},"ADD BUTTON FOR SELECTED DOCUMENT"),(0,i.kt)("p",null,"In below image we add ",(0,i.kt)("inlineCode",{parentName:"p"},"Button")," to open selected ",(0,i.kt)("em",{parentName:"p"},"SOLIDWORKS document"),"."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(69064).Z},(0,i.kt)("img",{alt:"add-button-for-solidworks-document",src:n(97698).Z,width:"1364",height:"726"}))),(0,i.kt)("p",null,(0,i.kt)("inlineCode",{parentName:"p"},"Button")," with ",(0,i.kt)("inlineCode",{parentName:"p"},"Content")," and ",(0,i.kt)("em",{parentName:"p"},"other properties")," are given below."),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-xml"},'<Button Width="350"\n        Height="50"\n        FontSize="18"\n        FontWeight="Medium"\n        Content="Open Solidworks" />\n')),(0,i.kt)("h2",{id:"update-viewmodel"},"UPDATE VIEWMODEL"),(0,i.kt)("p",null,"Now, we update our ",(0,i.kt)("inlineCode",{parentName:"p"},"MainWindowViewModel")," viewmodel, for showing data and adding functionalities."),(0,i.kt)("h3",{id:"add-documents-list"},"ADD DOCUMENTS LIST"),(0,i.kt)("p",null,"In below image we a list of SOLIDWORKS document in ",(0,i.kt)("inlineCode",{parentName:"p"},"MainWindowViewModel"),"."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(94569).Z},(0,i.kt)("img",{alt:"add-document-list-to-viewmodel",src:n(36974).Z,width:"1364",height:"726"}))),(0,i.kt)("p",null,"For this we use below code."),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-cs",metastring:"showLineNumbers",showLineNumbers:!0},"private ObservableCollection<string> _DocumentsList;\npublic ObservableCollection<string> DocumentsList\n{\n    get { return _DocumentsList; }\n    set { SetProperty(ref _DocumentsList, value); }\n}\n")),(0,i.kt)("p",null,"In above code, ",(0,i.kt)("inlineCode",{parentName:"p"},"_DocumentsList")," is private member of our ",(0,i.kt)("inlineCode",{parentName:"p"},"MainWindowViewModel")," class, whose value we set in the ",(0,i.kt)("inlineCode",{parentName:"p"},"Constructor")," of ",(0,i.kt)("inlineCode",{parentName:"p"},"MainWindowViewModel")," class."),(0,i.kt)("p",null,(0,i.kt)("inlineCode",{parentName:"p"},"DocumentsList")," will use for ",(0,i.kt)("inlineCode",{parentName:"p"},"Binding")," document list to our ",(0,i.kt)("inlineCode",{parentName:"p"},"ComboBox")," as ",(0,i.kt)("inlineCode",{parentName:"p"},"ItemSource"),"."),(0,i.kt)("p",null,"Here we use ",(0,i.kt)("inlineCode",{parentName:"p"},"ObservableCollection<T>")," because of ",(0,i.kt)("strong",{parentName:"p"},"MVVM"),"."),(0,i.kt)("p",null,"For more details please visit ",(0,i.kt)("strong",{parentName:"p"},(0,i.kt)("a",{parentName:"strong",href:"https://docs.microsoft.com/en-us/dotnet/api/system.collections.objectmodel.observablecollection-1?view=net-5.0"},"this link")),"."),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-cs",metastring:"showLineNumbers",showLineNumbers:!0},'public MainWindowViewModel()\n{\n    _DocumentsList = new ObservableCollection<string>\n    {\n        "Part Document",\n        "Assembly Document",\n        "Drawing Document"\n    };\n}\n')),(0,i.kt)("p",null,"In above code, we add SOLIDWORKS documents into our ",(0,i.kt)("inlineCode",{parentName:"p"},"_DocumentsList")," list."),(0,i.kt)("h3",{id:"binding-document-list-to-combobox"},"BINDING DOCUMENT LIST TO COMBOBOX"),(0,i.kt)("p",null,"In below image we ",(0,i.kt)("em",{parentName:"p"},"Bind")," our document list i.e. ",(0,i.kt)("inlineCode",{parentName:"p"},"DocumentsList")," to ",(0,i.kt)("inlineCode",{parentName:"p"},"ComboBox")," as ",(0,i.kt)("inlineCode",{parentName:"p"},"ItemSource"),"."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(22637).Z},(0,i.kt)("img",{alt:"binding-document-list",src:n(61052).Z,width:"1364",height:"726"}))),(0,i.kt)("p",null,"For Binding ",(0,i.kt)("inlineCode",{parentName:"p"},"DocumentsList")," we add following line."),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-xml"},'ItemsSource="{Binding DocumentsList}"\n')),(0,i.kt)("p",null,"After this update our ComboBox looks like as:"),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-xml"},'<ComboBox Width="350"\n          Height="30"\n          Margin="10"\n          VerticalAlignment="Center"\n          FontSize="16" \n          ItemsSource="{Binding DocumentsList}"/>\n')),(0,i.kt)("p",null,"When we ",(0,i.kt)("strong",{parentName:"p"},(0,i.kt)("inlineCode",{parentName:"strong"},"Run"))," our code, we get following window."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(41142).Z},(0,i.kt)("img",{alt:"comboBox-list-window",src:n(45011).Z,width:"1234",height:"657"}))),(0,i.kt)("p",null,"As I have mentioned in above image, if there are ",(0,i.kt)("strong",{parentName:"p"},"no item selected"),", we will get ",(0,i.kt)("strong",{parentName:"p"},"error"),' when we click "',(0,i.kt)("em",{parentName:"p"},"Open Solidworks"),'" button.'),(0,i.kt)("p",null,"To avoid this error we define ",(0,i.kt)("inlineCode",{parentName:"p"},"SelectedIndex")," property of ",(0,i.kt)("inlineCode",{parentName:"p"},"ComboBox")," to ",(0,i.kt)("strong",{parentName:"p"},"0"),"."),(0,i.kt)("p",null,"After this update our ComboBox looks like as:"),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-xml"},'<ComboBox Width="350"\n          Height="30"\n          Margin="10"\n          VerticalAlignment="Center"\n          FontSize="16" \n          SelectedIndex="0"\n          ItemsSource="{Binding DocumentsList}"/>\n')),(0,i.kt)("p",null,"When we ",(0,i.kt)("strong",{parentName:"p"},(0,i.kt)("inlineCode",{parentName:"strong"},"Run"))," our code, we get following window."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(24818).Z},(0,i.kt)("img",{alt:"part-document-selected-in-list",src:n(67714).Z,width:"1234",height:"657"}))),(0,i.kt)("h3",{id:"add-selected-value-in-combobox"},"ADD SELECTED VALUE IN COMBOBOX"),(0,i.kt)("p",null,"In our program, we want to open selected ",(0,i.kt)("em",{parentName:"p"},"SOLIDWORKS Document"),"."),(0,i.kt)("p",null,"To get the selected value, we need a property i.e. ",(0,i.kt)("inlineCode",{parentName:"p"},"SelectedDocument")," in our ",(0,i.kt)("inlineCode",{parentName:"p"},"MainWindowViewModel")," ViewModel and bind this property to ",(0,i.kt)("inlineCode",{parentName:"p"},"SelectValue")," property of ",(0,i.kt)("inlineCode",{parentName:"p"},"ComboBox"),"."),(0,i.kt)("p",null,"For more details please see below image."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(74372).Z},(0,i.kt)("img",{alt:"selectedvalue-binding",src:n(8715).Z,width:"1364",height:"726"}))),(0,i.kt)("h3",{id:"add-command-to-viewmodel"},"ADD COMMAND TO VIEWMODEL"),(0,i.kt)("p",null,"In our application to open selected SOLIDWORKS document, we need add a ",(0,i.kt)("em",{parentName:"p"},"Command")," to our button."),(0,i.kt)("p",null,"For this we need to do following:"),(0,i.kt)("ul",null,(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},"We need to create a ",(0,i.kt)("em",{parentName:"p"},"Prism Command")," i.e. ",(0,i.kt)("inlineCode",{parentName:"p"},"OpenSolidworksCommand")," in ",(0,i.kt)("inlineCode",{parentName:"p"},"MainWindowViewModel")," ViewModel.")),(0,i.kt)("li",{parentName:"ul"},(0,i.kt)("p",{parentName:"li"},"Bind this ",(0,i.kt)("inlineCode",{parentName:"p"},"OpenSolidworksCommand")," to our button."))),(0,i.kt)("p",null,"In below we see how to do this, also we checked the selected value."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(8080).Z},(0,i.kt)("img",{alt:"add-command-to-button",src:n(4593).Z,width:"1364",height:"726"}))),(0,i.kt)("h2",{id:"add-solidworks-references"},"ADD SOLIDWORKS REFERENCES"),(0,i.kt)("p",null,"For opening ",(0,i.kt)("strong",{parentName:"p"},"SOLIDWORKS")," we need to add some ",(0,i.kt)("em",{parentName:"p"},"references")," into our project."),(0,i.kt)("p",null,"Please see below image for how to add ",(0,i.kt)("em",{parentName:"p"},"SOLIDWORKS")," reference."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(68989).Z},(0,i.kt)("img",{alt:"add-solidworks-references",src:n(99268).Z,width:"1364",height:"726"}))),(0,i.kt)("h2",{id:"open-solidworks-document"},"OPEN SOLIDWORKS DOCUMENT"),(0,i.kt)("p",null,"Now for opening ",(0,i.kt)("em",{parentName:"p"},"SOLIDWORKS Document")," we need to add following code as shown in below image."),(0,i.kt)("pre",null,(0,i.kt)("code",{parentName:"pre",className:"language-cs",metastring:"showLineNumbers",showLineNumbers:!0},'void ExecuteOpenSolidworksCommand()\n{\n    // Create a new Instance of Solidworks Application\n    SldWorks.SldWorks swApp = new SldWorks.SldWorks();\n\n    // Make Solidworks visible\n    swApp.Visible = true;\n\n    // Variable to hold selected document\'s template path\n    string templatePath = string.Empty;\n\n    // Switch Conditional Statement\n    switch (SelectedDocument)\n    {\n        case "Part Document":\n            // Get default Part template path\n            templatePath = swApp.GetUserPreferenceStringValue((int)SwConst.swUserPreferenceStringValue_e.swDefaultTemplatePart);\n            break;\n        case "Assembly Document":\n            // Get default Assembly template path\n            templatePath = swApp.GetUserPreferenceStringValue((int)SwConst.swUserPreferenceStringValue_e.swDefaultTemplateAssembly);\n            break;\n        case "Drawing Document":\n            // Get default Drawing template path\n            templatePath = swApp.GetUserPreferenceStringValue((int)SwConst.swUserPreferenceStringValue_e.swDefaultTemplateDrawing);\n            break;\n    }\n\n    // Create a new Document as ModelDoc2 object\n    SldWorks.ModelDoc2 swDoc = swApp.NewDocument(templatePath, 0, 0, 0);\n}\n')),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(51183).Z},(0,i.kt)("img",{alt:"add-open-solidworks-document-code",src:n(51824).Z,width:"1364",height:"726"}))),(0,i.kt)("h2",{id:"final-result"},"FINAL RESULT"),(0,i.kt)("p",null,"Now, we have done everything needed to ",(0,i.kt)("strong",{parentName:"p"},"Open")," ",(0,i.kt)("em",{parentName:"p"},"SOLIDWORKS Document")," through our application."),(0,i.kt)("p",null,"Please see below image for final result of our work."),(0,i.kt)("p",null,(0,i.kt)("a",{target:"_blank",href:n(13630).Z},(0,i.kt)("img",{alt:"final-result",src:n(39589).Z,width:"1364",height:"726"}))),(0,i.kt)("hr",null),(0,i.kt)("p",null,(0,i.kt)("strong",{parentName:"p"},"This is it !!!")),(0,i.kt)("p",null,(0,i.kt)("em",{parentName:"p"},"I hope my efforts will helpful to someone!")),(0,i.kt)("p",null,"If you found anything to ",(0,i.kt)("strong",{parentName:"p"},"add or update"),", please let me know on my ",(0,i.kt)("em",{parentName:"p"},"e-mail"),"."),(0,i.kt)("p",null,"Hope this post helps you to ",(0,i.kt)("strong",{parentName:"p"},"Open SOLIDWORKS Documents")," from WPF PRISM Application."),(0,i.kt)("p",null,(0,i.kt)("em",{parentName:"p"},"If you like the post then please share it with your friends also.")),(0,i.kt)("p",null,(0,i.kt)("em",{parentName:"p"},"Do let me know by you like this post or not!")),(0,i.kt)("p",null,(0,i.kt)("em",{parentName:"p"},"Till then, Happy learning!!!")))}m.isMDXComponent=!0},44844:function(e,t,n){t.Z=n.p+"assets/files/1.create-new-project-04e3322a13f0f2ac34a218232853f5c9.svg"},1590:function(e,t,n){t.Z=n.p+"assets/files/1.create-open-solidworks-document-project-09b1be06b24b751d8377783fb89c8f09.gif"},8080:function(e,t,n){t.Z=n.p+"assets/files/10-add-command-to-button-edf78f8a18d6e3a80f03dc4104336649.gif"},68989:function(e,t,n){t.Z=n.p+"assets/files/11.add-solidworks-references-05111283bb25889c4d64541c2681169c.gif"},51183:function(e,t,n){t.Z=n.p+"assets/files/12.add-open-solidworks-document-code-22e57901167523f9efb729ef3817fe5c.gif"},64865:function(e,t,n){t.Z=n.p+"assets/files/2.build-solution-58c1a3710d12713d55b9470e86300e45.svg"},49851:function(e,t,n){t.Z=n.p+"assets/files/2.update-window-startup-location-width-height-a26562a85ce89ecbaae63b6cb1380fe5.gif"},88936:function(e,t,n){t.Z=n.p+"assets/files/3.before-build-solution-47bc2cf0209c7667a480a1c532c9e3d8.svg"},56089:function(e,t,n){t.Z=n.p+"assets/files/3.remove-content-control-777c7d107c145dabdbaa35717dae070a.gif"},30166:function(e,t,n){t.Z=n.p+"assets/files/4.add-textblock-0f325555575a8c39fbee4b8172294022.gif"},34255:function(e,t,n){t.Z=n.p+"assets/files/4.after-build-solution-0d9704d06704b72f1f4d760578c252de.svg"},4527:function(e,t,n){t.Z=n.p+"assets/files/5.add-documents-combobox-85c0117e9948c9a26b9f2cbc927bdf44.gif"},41142:function(e,t,n){t.Z=n.p+"assets/files/5.comboBox-list-window-ca19c4d045197375c71f1d02d0910e08.svg"},69064:function(e,t,n){t.Z=n.p+"assets/files/6.add-button-for-solidworks-document-5c9f784af4b68f1b043b6c69ad781f64.gif"},24818:function(e,t,n){t.Z=n.p+"assets/files/6.part-document-selected-in-list-a39df3867f92f605e33b4df10697c8e5.svg"},94569:function(e,t,n){t.Z=n.p+"assets/files/7.add-document-list-to-viewmodel-8729a77785d74ad00d35135eca07e375.gif"},22637:function(e,t,n){t.Z=n.p+"assets/files/8.binding-document-list-fab3a636bc61fbc1c3499c1ecfdd1fc9.gif"},74372:function(e,t,n){t.Z=n.p+"assets/files/9.selectedvalue-binding-ee728082901b1fbbffe92941d3b6496a.gif"},13630:function(e,t,n){t.Z=n.p+"assets/files/final-result-67aa12a239288f135d7f605e4842b479.gif"},37514:function(e,t,n){t.Z=n.p+"assets/images/1.create-new-project-04e3322a13f0f2ac34a218232853f5c9.svg"},35334:function(e,t,n){t.Z=n.p+"assets/images/1.create-open-solidworks-document-project-09b1be06b24b751d8377783fb89c8f09.gif"},4593:function(e,t,n){t.Z=n.p+"assets/images/10-add-command-to-button-edf78f8a18d6e3a80f03dc4104336649.gif"},99268:function(e,t,n){t.Z=n.p+"assets/images/11.add-solidworks-references-05111283bb25889c4d64541c2681169c.gif"},51824:function(e,t,n){t.Z=n.p+"assets/images/12.add-open-solidworks-document-code-22e57901167523f9efb729ef3817fe5c.gif"},76216:function(e,t,n){t.Z=n.p+"assets/images/2.build-solution-58c1a3710d12713d55b9470e86300e45.svg"},74608:function(e,t,n){t.Z=n.p+"assets/images/2.update-window-startup-location-width-height-a26562a85ce89ecbaae63b6cb1380fe5.gif"},53095:function(e,t,n){t.Z=n.p+"assets/images/3.before-build-solution-47bc2cf0209c7667a480a1c532c9e3d8.svg"},16701:function(e,t,n){t.Z=n.p+"assets/images/3.remove-content-control-777c7d107c145dabdbaa35717dae070a.gif"},66862:function(e,t,n){t.Z=n.p+"assets/images/4.add-textblock-0f325555575a8c39fbee4b8172294022.gif"},59496:function(e,t,n){t.Z=n.p+"assets/images/4.after-build-solution-0d9704d06704b72f1f4d760578c252de.svg"},56342:function(e,t,n){t.Z=n.p+"assets/images/5.add-documents-combobox-85c0117e9948c9a26b9f2cbc927bdf44.gif"},45011:function(e,t,n){t.Z=n.p+"assets/images/5.comboBox-list-window-ca19c4d045197375c71f1d02d0910e08.svg"},97698:function(e,t,n){t.Z=n.p+"assets/images/6.add-button-for-solidworks-document-5c9f784af4b68f1b043b6c69ad781f64.gif"},67714:function(e,t,n){t.Z=n.p+"assets/images/6.part-document-selected-in-list-a39df3867f92f605e33b4df10697c8e5.svg"},36974:function(e,t,n){t.Z=n.p+"assets/images/7.add-document-list-to-viewmodel-8729a77785d74ad00d35135eca07e375.gif"},61052:function(e,t,n){t.Z=n.p+"assets/images/8.binding-document-list-fab3a636bc61fbc1c3499c1ecfdd1fc9.gif"},8715:function(e,t,n){t.Z=n.p+"assets/images/9.selectedvalue-binding-ee728082901b1fbbffe92941d3b6496a.gif"},39589:function(e,t,n){t.Z=n.p+"assets/images/final-result-67aa12a239288f135d7f605e4842b479.gif"}}]);