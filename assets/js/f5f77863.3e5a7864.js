"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[3901],{3905:(e,t,n)=>{n.d(t,{Zo:()=>p,kt:()=>c});var a=n(67294);function l(e,t,n){return t in e?Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}):e[t]=n,e}function r(e,t){var n=Object.keys(e);if(Object.getOwnPropertySymbols){var a=Object.getOwnPropertySymbols(e);t&&(a=a.filter((function(t){return Object.getOwnPropertyDescriptor(e,t).enumerable}))),n.push.apply(n,a)}return n}function i(e){for(var t=1;t<arguments.length;t++){var n=null!=arguments[t]?arguments[t]:{};t%2?r(Object(n),!0).forEach((function(t){l(e,t,n[t])})):Object.getOwnPropertyDescriptors?Object.defineProperties(e,Object.getOwnPropertyDescriptors(n)):r(Object(n)).forEach((function(t){Object.defineProperty(e,t,Object.getOwnPropertyDescriptor(n,t))}))}return e}function o(e,t){if(null==e)return{};var n,a,l=function(e,t){if(null==e)return{};var n,a,l={},r=Object.keys(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||(l[n]=e[n]);return l}(e,t);if(Object.getOwnPropertySymbols){var r=Object.getOwnPropertySymbols(e);for(a=0;a<r.length;a++)n=r[a],t.indexOf(n)>=0||Object.prototype.propertyIsEnumerable.call(e,n)&&(l[n]=e[n])}return l}var s=a.createContext({}),u=function(e){var t=a.useContext(s),n=t;return e&&(n="function"==typeof e?e(t):i(i({},t),e)),n},p=function(e){var t=u(e.components);return a.createElement(s.Provider,{value:t},e.children)},m={inlineCode:"code",wrapper:function(e){var t=e.children;return a.createElement(a.Fragment,{},t)}},d=a.forwardRef((function(e,t){var n=e.components,l=e.mdxType,r=e.originalType,s=e.parentName,p=o(e,["components","mdxType","originalType","parentName"]),d=u(n),c=l,g=d["".concat(s,".").concat(c)]||d[c]||m[c]||r;return n?a.createElement(g,i(i({ref:t},p),{},{components:n})):a.createElement(g,i({ref:t},p))}));function c(e,t){var n=arguments,l=t&&t.mdxType;if("string"==typeof e||l){var r=n.length,i=new Array(r);i[0]=d;var o={};for(var s in t)hasOwnProperty.call(t,s)&&(o[s]=t[s]);o.originalType=e,o.mdxType="string"==typeof e?e:l,i[1]=o;for(var u=2;u<r;u++)i[u]=n[u];return a.createElement.apply(null,i)}return a.createElement.apply(null,n)}d.displayName="MDXCreateElement"},93202:(e,t,n)=>{n.r(t),n.d(t,{assets:()=>s,contentTitle:()=>i,default:()=>m,frontMatter:()=>r,metadata:()=>o,toc:()=>u});var a=n(87462),l=(n(67294),n(3905));const r={title:"VBA MsgBox Function",tags:["VBA"],permalink:"/vba/msgBox-function/"},i=void 0,o={unversionedId:"vba-msgBox-function",id:"vba-msgBox-function",title:"VBA MsgBox Function",description:"You\u2019re probably already familiar with the VBA MsgBox function \u2014 I use it quite a bit in the examples.",source:"@site/docs/vba/24-vba-msgBox-function.md",sourceDirName:".",slug:"/vba-msgBox-function",permalink:"/vba/vba-msgBox-function",draft:!1,tags:[{label:"VBA",permalink:"/vba/tags/vba"}],version:"current",sidebarPosition:24,frontMatter:{title:"VBA MsgBox Function",tags:["VBA"],permalink:"/vba/msgBox-function/"},sidebar:"tutorialSidebar",previous:{title:"VBA Dialog Boxes",permalink:"/vba/vba-dialog-boxes"},next:{title:"VBA InputBox Function",permalink:"/vba/vba-inputbox-function"}},s={},u=[{value:"Getting a response from a message box",id:"getting-a-response-from-a-message-box",level:2},{value:"Customizing message boxes",id:"customizing-message-boxes",level:2}],p={toc:u};function m(e){let{components:t,...r}=e;return(0,l.kt)("wrapper",(0,a.Z)({},p,r,{components:t,mdxType:"MDXLayout"}),(0,l.kt)("p",null,"You\u2019re probably already familiar with the VBA ",(0,l.kt)("inlineCode",{parentName:"p"},"MsgBox")," function \u2014 I use it quite a bit in the examples. "),(0,l.kt)("p",null,"The ",(0,l.kt)("inlineCode",{parentName:"p"},"MsgBox")," function, which accepts the arguments shown in below table, is handy for displaying information and getting simple user input. "),(0,l.kt)("p",null,"It\u2019s able to get user input because it\u2019s a function. "),(0,l.kt)("p",null,"A ",(0,l.kt)("em",{parentName:"p"},"function"),", as you recall, returns a value. "),(0,l.kt)("p",null,"In the case of the ",(0,l.kt)("inlineCode",{parentName:"p"},"Msgbox")," function, it uses a dialog box to get the value that it returns. "),(0,l.kt)("p",null,"Keep reading to see exactly how it works."),(0,l.kt)("p",null,"Here\u2019s a simplified version of the syntax for the MsgBox function:"),(0,l.kt)("pre",null,(0,l.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' MsgBox Structure\nMsgBox(prompt[, buttons][, title])\n")),(0,l.kt)("table",null,(0,l.kt)("thead",{parentName:"table"},(0,l.kt)("tr",{parentName:"thead"},(0,l.kt)("th",{parentName:"tr",align:null},"Arguments"),(0,l.kt)("th",{parentName:"tr",align:null},"What it does"))),(0,l.kt)("tbody",{parentName:"table"},(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"prompt"),(0,l.kt)("td",{parentName:"tr",align:null},"The text your application displays in the message box")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"buttons"),(0,l.kt)("td",{parentName:"tr",align:null},"A number that specifies which buttons (along with what icon) appear in the message box (optional)")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"title"),(0,l.kt)("td",{parentName:"tr",align:null},"The text that appears in the message box\u2019s title bar (optional) displaying a simple message box")))),(0,l.kt)("p",null,"You can use the ",(0,l.kt)("em",{parentName:"p"},"MsgBox")," function in two ways:"),(0,l.kt)("ul",null,(0,l.kt)("li",{parentName:"ul"},"To simply show a message to the user. In this case, you don\u2019t care about the result returned by the function."),(0,l.kt)("li",{parentName:"ul"},"To get a response from the user. In this case, you do care about the result returned by the function. The result depends on the button that the user clicks.")),(0,l.kt)("p",null,"If you use the ",(0,l.kt)("em",{parentName:"p"},"MsgBox")," function by itself, don\u2019t include parentheses around the arguments. "),(0,l.kt)("p",null,"The following example simply displays a message and does not return a result. "),(0,l.kt)("p",null,"When the message is displayed, the code stops until the user clicks ",(0,l.kt)("inlineCode",{parentName:"p"},"OK"),"."),(0,l.kt)("pre",null,(0,l.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' MsgBox function Example\nSub main()\n  MsgBox "Hello, world!"\nEnd Sub\n')),(0,l.kt)("p",null,"Below figure shows how this message box looks:"),(0,l.kt)("p",null,(0,l.kt)("img",{alt:"A-Simple-Message-Box",src:n(89938).Z,width:"145",height:"153"})),(0,l.kt)("h2",{id:"getting-a-response-from-a-message-box"},"Getting a response from a message box"),(0,l.kt)("p",null,"If you display a message box that has more than just an ",(0,l.kt)("strong",{parentName:"p"},"OK")," button, you\u2019ll probably want to know which button the user clicks. "),(0,l.kt)("p",null,"The ",(0,l.kt)("em",{parentName:"p"},"MsgBox function")," can return a value that represents which button is clicked. "),(0,l.kt)("p",null,"You can assign the result of the MsgBox function to a variable."),(0,l.kt)("p",null,"In the following code, I use some built-in constants that make it easy to work with the values returned by ",(0,l.kt)("inlineCode",{parentName:"p"},"MsgBox"),":"),(0,l.kt)("pre",null,(0,l.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' MsgBox built-in constants Example\nSub GetAnswer()\n  Dim Ans as Integer\n  Ans = MsgBox (\"Did you eat lunch?\", vbYesNo)\n  Select Case Ans\n    Case vbYes\n    '......[Some code here]....\n    Case vbNo\n    '......[Some code here]....\n  End Select\nEnd Sub\n")),(0,l.kt)("p",null,"Below figure shows how it looks. "),(0,l.kt)("p",null,"When you execute this procedure, the ",(0,l.kt)("inlineCode",{parentName:"p"},"Ans")," variable is assigned a value of either ",(0,l.kt)("inlineCode",{parentName:"p"},"vbYes")," or ",(0,l.kt)("inlineCode",{parentName:"p"},"vbNo"),", depending on which button the user clicks. "),(0,l.kt)("p",null,"The ",(0,l.kt)("inlineCode",{parentName:"p"},"Select")," Case statement uses the ",(0,l.kt)("inlineCode",{parentName:"p"},"Ans")," value to determine which action the code should perform."),(0,l.kt)("p",null,(0,l.kt)("img",{alt:"A-Simple-Message-Box-with-two-buttons",src:n(92785).Z,width:"239",height:"150"})),(0,l.kt)("p",null,"You can also use the ",(0,l.kt)("inlineCode",{parentName:"p"},"MsgBox")," function result without using a variable, as the following example demonstrates:"),(0,l.kt)("pre",null,(0,l.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' MsgBox without variable\nSub GetAnswer2()\n  If MsgBox (\"Continue?\", vbYesNo) = vbYes Then\n  '......[Some code here]....\n  Else\n  '......[Some code here]....\n  End If\nEnd Sub\n")),(0,l.kt)("h2",{id:"customizing-message-boxes"},"Customizing message boxes"),(0,l.kt)("p",null,"The flexibility of the buttons argument makes it easy to customize your message boxes. "),(0,l.kt)("p",null,"You can specify which buttons to display, determine whether an icon appears, and decide which button is the default (the default button is \u201cclicked\u201d if the user presses ",(0,l.kt)("inlineCode",{parentName:"p"},"Enter"),")."),(0,l.kt)("p",null,"Below table lists some of the built-in constants you can use for the buttons argument. "),(0,l.kt)("p",null,"If you prefer, you can use the value rather than a constant (but I think using the built-in constants is a lot easier)."),(0,l.kt)("table",null,(0,l.kt)("thead",{parentName:"table"},(0,l.kt)("tr",{parentName:"thead"},(0,l.kt)("th",{parentName:"tr",align:null},"Constant"),(0,l.kt)("th",{parentName:"tr",align:null},"Value"),(0,l.kt)("th",{parentName:"tr",align:null},"What it does"))),(0,l.kt)("tbody",{parentName:"table"},(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbOKOnly"),(0,l.kt)("td",{parentName:"tr",align:null},"0"),(0,l.kt)("td",{parentName:"tr",align:null},"Display OK button only.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbOKCancel"),(0,l.kt)("td",{parentName:"tr",align:null},"1"),(0,l.kt)("td",{parentName:"tr",align:null},"Display OK and Cancel buttons")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbAbortRetryIgnore"),(0,l.kt)("td",{parentName:"tr",align:null},"2"),(0,l.kt)("td",{parentName:"tr",align:null},"Displays Abort, Retry, and Ignore buttons.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbYesNoCancel"),(0,l.kt)("td",{parentName:"tr",align:null},"3"),(0,l.kt)("td",{parentName:"tr",align:null},"Displays Yes, No, and Cancel buttons.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbYesNo"),(0,l.kt)("td",{parentName:"tr",align:null},"4"),(0,l.kt)("td",{parentName:"tr",align:null},"Displays Yes and No buttons.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbRetryCancel"),(0,l.kt)("td",{parentName:"tr",align:null},"5"),(0,l.kt)("td",{parentName:"tr",align:null},"Displays Retry and Cancel buttons.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbCritical"),(0,l.kt)("td",{parentName:"tr",align:null},"16"),(0,l.kt)("td",{parentName:"tr",align:null},"Displays Critical Message icon.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbQuestion"),(0,l.kt)("td",{parentName:"tr",align:null},"32"),(0,l.kt)("td",{parentName:"tr",align:null},"Displays Warning Query icon.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbExclamation"),(0,l.kt)("td",{parentName:"tr",align:null},"48"),(0,l.kt)("td",{parentName:"tr",align:null},"Displays Warning Message icon.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbInformation"),(0,l.kt)("td",{parentName:"tr",align:null},"64"),(0,l.kt)("td",{parentName:"tr",align:null},"Displays Information Message icon.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbDefaultButton1"),(0,l.kt)("td",{parentName:"tr",align:null},"0"),(0,l.kt)("td",{parentName:"tr",align:null},"First button is default.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbDefaultButton2"),(0,l.kt)("td",{parentName:"tr",align:null},"256"),(0,l.kt)("td",{parentName:"tr",align:null},"Second button is default.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbDefaultButton3"),(0,l.kt)("td",{parentName:"tr",align:null},"512"),(0,l.kt)("td",{parentName:"tr",align:null},"Third button is default.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbDefaultButton4"),(0,l.kt)("td",{parentName:"tr",align:null},"768"),(0,l.kt)("td",{parentName:"tr",align:null},"Fourth button is default.")))),(0,l.kt)("p",null,"For using more than one of these constants as an argument, just connect them with a ",(0,l.kt)("inlineCode",{parentName:"p"},"+")," operator. "),(0,l.kt)("p",null,"For example, to display a message box with ",(0,l.kt)("inlineCode",{parentName:"p"},"Yes")," and ",(0,l.kt)("inlineCode",{parentName:"p"},"No")," buttons and an exclamation icon, use the following expression as the second ",(0,l.kt)("em",{parentName:"p"},"MsgBox")," argument:"),(0,l.kt)("pre",null,(0,l.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},"' Using multiple MsgBox built-in constants\nvbYesNo + vbExclamation\n")),(0,l.kt)("p",null,"Or, if you prefer to make your code less understandable, use a value of ",(0,l.kt)("em",{parentName:"p"},"52 (that is, 4 + 48)"),"."),(0,l.kt)("p",null,"The following example uses a combination of constants to display a message box with a ",(0,l.kt)("inlineCode",{parentName:"p"},"Yes button")," and a ",(0,l.kt)("inlineCode",{parentName:"p"},"No button")," (",(0,l.kt)("inlineCode",{parentName:"p"},"vbYesNo"),") as well as a question mark icon (",(0,l.kt)("inlineCode",{parentName:"p"},"vbQuestion"),"). "),(0,l.kt)("p",null,"The constant ",(0,l.kt)("inlineCode",{parentName:"p"},"vbDefaultButton2")," designates the second button (",(0,l.kt)("inlineCode",{parentName:"p"},"No"),") as the default button \u2014 that is, the button that is clicked if the user presses ",(0,l.kt)("inlineCode",{parentName:"p"},"Enter"),". "),(0,l.kt)("p",null,"For simplicity, we assign these constants to the ",(0,l.kt)("inlineCode",{parentName:"p"},"Config")," variable and then use ",(0,l.kt)("inlineCode",{parentName:"p"},"Config")," as the second argument in the ",(0,l.kt)("em",{parentName:"p"},"MsgBox")," function:"),(0,l.kt)("pre",null,(0,l.kt)("code",{parentName:"pre",className:"language-vb",metastring:"showlinenumbers showLineNumbers",showlinenumbers:!0,showLineNumbers:!0},'\' Using multiple MsgBox built-in constants\nSub GetAnswer3()\n  Dim Config As Integer\n  Dim Ans as Integer\n  Config = vbYesNo + vbQuestion + vbDefaultButton2\n  Ans = MsgBox("Is part opened?", Config)\n  If Ans = vbYes Then OpenPart\nEnd Sub\n')),(0,l.kt)("p",null,"Below figure shows the message box application displays when you execute the ",(0,l.kt)("inlineCode",{parentName:"p"},"GetAnswer3")," procedure. "),(0,l.kt)("p",null,"If the user clicks the ",(0,l.kt)("em",{parentName:"p"},"Yes button"),", the routine executes the procedure named ",(0,l.kt)("inlineCode",{parentName:"p"},"OpenPart")," (which is not shown). "),(0,l.kt)("p",null,"If the user clicks the ",(0,l.kt)("em",{parentName:"p"},"No button")," (or presses ",(0,l.kt)("inlineCode",{parentName:"p"},"Enter"),"), the routine ends with no action. "),(0,l.kt)("p",null,"Because I omitted the title argument in the ",(0,l.kt)("em",{parentName:"p"},"MsgBox")," function, our application uses the default title, in my case it is ",(0,l.kt)("em",{parentName:"p"},"Solidworks"),"."),(0,l.kt)("p",null,(0,l.kt)("img",{alt:"MsgBox-function-button",src:n(12200).Z,width:"241",height:"170"})),(0,l.kt)("p",null,"Previous examples have used constants (such as ",(0,l.kt)("inlineCode",{parentName:"p"},"vbYes")," and ",(0,l.kt)("inlineCode",{parentName:"p"},"vbNo"),") for the return value of a ",(0,l.kt)("em",{parentName:"p"},"MsgBox")," function. "),(0,l.kt)("p",null,"Besides these two constants, below table lists a few others."),(0,l.kt)("table",null,(0,l.kt)("thead",{parentName:"table"},(0,l.kt)("tr",{parentName:"thead"},(0,l.kt)("th",{parentName:"tr",align:null},"Constant"),(0,l.kt)("th",{parentName:"tr",align:null},"Value"),(0,l.kt)("th",{parentName:"tr",align:null},"What it does"))),(0,l.kt)("tbody",{parentName:"table"},(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbOK"),(0,l.kt)("td",{parentName:"tr",align:null},"1"),(0,l.kt)("td",{parentName:"tr",align:null},"User clicked OK.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbCancel"),(0,l.kt)("td",{parentName:"tr",align:null},"2"),(0,l.kt)("td",{parentName:"tr",align:null},"User clicked Cancel.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbAbort"),(0,l.kt)("td",{parentName:"tr",align:null},"3"),(0,l.kt)("td",{parentName:"tr",align:null},"User clicked Abort.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbRetry"),(0,l.kt)("td",{parentName:"tr",align:null},"4"),(0,l.kt)("td",{parentName:"tr",align:null},"User clicked Retry.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbIgnore"),(0,l.kt)("td",{parentName:"tr",align:null},"5"),(0,l.kt)("td",{parentName:"tr",align:null},"User clicked Ignore.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbYes"),(0,l.kt)("td",{parentName:"tr",align:null},"6"),(0,l.kt)("td",{parentName:"tr",align:null},"User clicked Yes.")),(0,l.kt)("tr",{parentName:"tbody"},(0,l.kt)("td",{parentName:"tr",align:null},"vbNo"),(0,l.kt)("td",{parentName:"tr",align:null},"7"),(0,l.kt)("td",{parentName:"tr",align:null},"User clicked No.")))),(0,l.kt)("p",null,"Next post will be about ",(0,l.kt)("strong",{parentName:"p"},(0,l.kt)("em",{parentName:"strong"},"VBA InputBox Function")),"."))}m.isMDXComponent=!0},89938:(e,t,n)=>{n.d(t,{Z:()=>a});const a="data:image/png;base64,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"},92785:(e,t,n)=>{n.d(t,{Z:()=>a});const a="data:image/png;base64,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"},12200:(e,t,n)=>{n.d(t,{Z:()=>a});const a="data:image/png;base64,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"}}]);