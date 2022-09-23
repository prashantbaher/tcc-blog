"use strict";(self.webpackChunkdocs_website=self.webpackChunkdocs_website||[]).push([[3461],{3905:function(t,n,e){e.d(n,{Zo:function(){return d},kt:function(){return f}});var l=e(67294);function r(t,n,e){return n in t?Object.defineProperty(t,n,{value:e,enumerable:!0,configurable:!0,writable:!0}):t[n]=e,t}function u(t,n){var e=Object.keys(t);if(Object.getOwnPropertySymbols){var l=Object.getOwnPropertySymbols(t);n&&(l=l.filter((function(n){return Object.getOwnPropertyDescriptor(t,n).enumerable}))),e.push.apply(e,l)}return e}function a(t){for(var n=1;n<arguments.length;n++){var e=null!=arguments[n]?arguments[n]:{};n%2?u(Object(e),!0).forEach((function(n){r(t,n,e[n])})):Object.getOwnPropertyDescriptors?Object.defineProperties(t,Object.getOwnPropertyDescriptors(e)):u(Object(e)).forEach((function(n){Object.defineProperty(t,n,Object.getOwnPropertyDescriptor(e,n))}))}return t}function o(t,n){if(null==t)return{};var e,l,r=function(t,n){if(null==t)return{};var e,l,r={},u=Object.keys(t);for(l=0;l<u.length;l++)e=u[l],n.indexOf(e)>=0||(r[e]=t[e]);return r}(t,n);if(Object.getOwnPropertySymbols){var u=Object.getOwnPropertySymbols(t);for(l=0;l<u.length;l++)e=u[l],n.indexOf(e)>=0||Object.prototype.propertyIsEnumerable.call(t,e)&&(r[e]=t[e])}return r}var i=l.createContext({}),s=function(t){var n=l.useContext(i),e=n;return t&&(e="function"==typeof t?t(n):a(a({},n),t)),e},d=function(t){var n=s(t.components);return l.createElement(i.Provider,{value:n},t.children)},k={inlineCode:"code",wrapper:function(t){var n=t.children;return l.createElement(l.Fragment,{},n)}},c=l.forwardRef((function(t,n){var e=t.components,r=t.mdxType,u=t.originalType,i=t.parentName,d=o(t,["components","mdxType","originalType","parentName"]),c=s(e),f=r,m=c["".concat(i,".").concat(f)]||c[f]||k[f]||u;return e?l.createElement(m,a(a({ref:n},d),{},{components:e})):l.createElement(m,a({ref:n},d))}));function f(t,n){var e=arguments,r=n&&n.mdxType;if("string"==typeof t||r){var u=e.length,a=new Array(u);a[0]=c;var o={};for(var i in n)hasOwnProperty.call(n,i)&&(o[i]=n[i]);o.originalType=t,o.mdxType="string"==typeof t?t:r,a[1]=o;for(var s=2;s<u;s++)a[s]=e[s];return l.createElement.apply(null,a)}return l.createElement.apply(null,e)}c.displayName="MDXCreateElement"},76855:function(t,n,e){e.r(n),e.d(n,{assets:function(){return d},contentTitle:function(){return i},default:function(){return f},frontMatter:function(){return o},metadata:function(){return s},toc:function(){return k}});var l=e(87462),r=e(63366),u=(e(67294),e(3905)),a=["components"],o={title:"VBA Functions that do more",tags:["VBA"],permalink:"/vba/more-functions/"},i=void 0,s={unversionedId:"vba-more-function",id:"vba-more-function",title:"VBA Functions that do more",description:"A few VBA functions go above and beyond the call of duty. Rather than simply return a value, these functions have some useful side effects.",source:"@site/docs/vba/16-vba-more-function.md",sourceDirName:".",slug:"/vba-more-function",permalink:"/vba/vba-more-function",draft:!1,tags:[{label:"VBA",permalink:"/vba/tags/vba"}],version:"current",sidebarPosition:16,frontMatter:{title:"VBA Functions that do more",tags:["VBA"],permalink:"/vba/more-functions/"},sidebar:"tutorialSidebar",previous:{title:"VBA Functions",permalink:"/vba/vba-functions"},next:{title:"Controlling Program Flow and Making Decisions",permalink:"/vba/vba-controlling-flow-making-desicions"}},d={},k=[{value:"Discovering VBA functions",id:"discovering-vba-functions",level:2}],c={toc:k};function f(t){var n=t.components,e=(0,r.Z)(t,a);return(0,u.kt)("wrapper",(0,l.Z)({},c,e,{components:n,mdxType:"MDXLayout"}),(0,u.kt)("p",null,"A few VBA ",(0,u.kt)("inlineCode",{parentName:"p"},"functions")," go above and beyond the call of duty. Rather than simply return a value, these functions have some useful side effects. "),(0,u.kt)("p",null,"Below table lists them."),(0,u.kt)("table",{class:"w3-table-all w3-mobile w3-card-4"},(0,u.kt)("tr",null,(0,u.kt)("th",{class:"w3-center",colspan:"2"},"Functions with Useful Side Benefits")),(0,u.kt)("tr",null,(0,u.kt)("th",null,"Function"),(0,u.kt)("th",null,"What is does")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"MsgBox"),(0,u.kt)("td",null,"Displays a handy dialog box containing a message and buttons. The function returns a code that identifies which button the user clicks.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"InputBox"),(0,u.kt)("td",null,"Displays a simple dialog box that asks the user for some input. The function returns whatever the user enters into the dialog box.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Shell"),(0,u.kt)("td",null,"Executes another program. The function returns the task ID (a unique identifier) of the other program (or an error if the function can\u2019t start the other program).")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"GetObject/CreateObject"),(0,u.kt)("td",null,"Returns/Create a reference to an object provided by an ActiveX component. (If you don't understand, don't bother about it. Just remember we use this function to for checking & creating objects in later topics)"))),(0,u.kt)("hr",null),(0,u.kt)("h2",{id:"discovering-vba-functions"},"Discovering VBA functions"),(0,u.kt)("p",null,"How do we find out which function does VBA provides? "),(0,u.kt)("p",null,"The best source is the ",(0,u.kt)("em",{parentName:"p"},"Visual Basic Help system")," in build in your CAD Application. "),(0,u.kt)("p",null,"I compiled a partial list of ",(0,u.kt)("inlineCode",{parentName:"p"},"functions"),", which I share with you in following Table. "),(0,u.kt)("p",null,"I omitted some of the more specialized or obscure functions. "),(0,u.kt)("p",null,"For complete details on a particular function, type the function name into a VBA module, move the cursor anywhere in the text, and press ",(0,u.kt)("inlineCode",{parentName:"p"},"F1"),". "),(0,u.kt)("table",{class:"w3-table-all w3-mobile w3-card-4"},(0,u.kt)("tr",null,(0,u.kt)("th",{class:"w3-center",colspan:"2"},"VBA\u2019s Most Useful Built-In Functions")),(0,u.kt)("tr",null,(0,u.kt)("th",null,"Function"),(0,u.kt)("th",null,"What is does")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Abs"),(0,u.kt)("td",null,"Returns a number\u2019s absolute value.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Array"),(0,u.kt)("td",null,"Returns a variant containing an array.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Asc"),(0,u.kt)("td",null,"Converts the first character of a string to its ASCII value.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Atn"),(0,u.kt)("td",null,"Returns the arctangent of a number.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Choose"),(0,u.kt)("td",null,"Returns a value from a list of items.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Chr"),(0,u.kt)("td",null,"Converts an ANSI value to a string.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Cos"),(0,u.kt)("td",null,"Returns a number\u2019s cosine.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"CurDir"),(0,u.kt)("td",null,"Returns the current path.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Date"),(0,u.kt)("td",null,"Returns the current system date.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"DateAdd"),(0,u.kt)("td",null,"Returns a date to which a specified time interval has been added \u2014 for example, one month from a particular date.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"DatePart"),(0,u.kt)("td",null,"Returns an integer containing the specified part of a given date \u2014 for example, a date\u2019s day of the year.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"DateSerial"),(0,u.kt)("td",null,"Converts a date to a serial number.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"DateValue"),(0,u.kt)("td",null,"Converts a string to a date.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Day"),(0,u.kt)("td",null,"Returns the day of the month from a date value.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Dir"),(0,u.kt)("td",null,"Returns the name of a file or directory that matches a pattern.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Erl"),(0,u.kt)("td",null,"Returns the line number that caused an error.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Err"),(0,u.kt)("td",null,"Returns the error number of an error condition.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Error"),(0,u.kt)("td",null,"Returns the error message that corresponds to an error number.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Exp"),(0,u.kt)("td",null,"Returns the base of the natural logarithm (e) raised to a power.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"FileLen"),(0,u.kt)("td",null,"Returns the number of bytes in a file.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Fix"),(0,u.kt)("td",null,"Returns a number\u2019s integer portion.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Format"),(0,u.kt)("td",null,"Displays an expression in a particular format.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"GetSetting"),(0,u.kt)("td",null,"Returns a value from the Windows registry.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Hex"),(0,u.kt)("td",null,"Converts from decimal to hexadecimal.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Hour"),(0,u.kt)("td",null,"Returns the hours portion of a time.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"InputBox"),(0,u.kt)("td",null,"Displays a box to prompt a user for input.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"InStr"),(0,u.kt)("td",null,"Returns the position of a string within another string.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Int"),(0,u.kt)("td",null,"Returns the integer portion of a number.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"IPmt"),(0,u.kt)("td",null,"Returns the interest payment for an annuity or loan.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"IsArray"),(0,u.kt)("td",null,"Returns True if a variable is an array.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"IsDate"),(0,u.kt)("td",null,"Returns True if an expression is a date.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"IsEmpty"),(0,u.kt)("td",null,"Returns True if a variable has not been initialized.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"IsError"),(0,u.kt)("td",null,"Returns True if an expression is an error value.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"IsMissing"),(0,u.kt)("td",null,"Returns True if an optional argument was not passed to a procedure.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"IsNull"),(0,u.kt)("td",null,"Returns True if an expression contains no valid data.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"IsNumeric"),(0,u.kt)("td",null,"Returns True if an expression can be evaluated as a number.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"IsObject"),(0,u.kt)("td",null,"Returns True if an expression references an OLE Automation object.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"LBound"),(0,u.kt)("td",null,"Returns the smallest subscript for a dimension of an array.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"LCase"),(0,u.kt)("td",null,"Returns a string converted to lowercase.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Left"),(0,u.kt)("td",null,"Returns a specified number of characters from the left of a string.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Len"),(0,u.kt)("td",null,"Returns the number of characters in a string.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Log"),(0,u.kt)("td",null,"Returns the natural logarithm of a number to base.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"LTrim"),(0,u.kt)("td",null,"Returns a copy of a string, with any leading spaces removed.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Mid"),(0,u.kt)("td",null,"Returns a specified number of characters from a string.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Minutes"),(0,u.kt)("td",null,"Returns the minutes portion of a time value.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Month"),(0,u.kt)("td",null,"Returns the month from a date value.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"MsgBox"),(0,u.kt)("td",null,"Displays a message box and (optionally) returns a value.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Now"),(0,u.kt)("td",null,"Returns the current system date and time.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"RGB"),(0,u.kt)("td",null,"Returns a numeric RGB value representing a color.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Replace"),(0,u.kt)("td",null,"Replaces a substring in a string with another substring.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Right"),(0,u.kt)("td",null,"Returns a specified number of characters from the right of a string.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Rnd"),(0,u.kt)("td",null,"Returns a random number between 0 and 1.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"RTrim"),(0,u.kt)("td",null,"Returns a copy of a string, with any trailing spaces removed.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Second"),(0,u.kt)("td",null,"Returns the seconds portion of a time value.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Sgn"),(0,u.kt)("td",null,"Returns an integer that indicates a number\u2019s sign.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Shell"),(0,u.kt)("td",null,"Runs an executable program.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Sin"),(0,u.kt)("td",null,"Returns a number\u2019s sine.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Space"),(0,u.kt)("td",null,"Returns a string with a specified number of spaces.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Split"),(0,u.kt)("td",null,"Splits a string into parts, using a delimiting character.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Sqr"),(0,u.kt)("td",null,"Returns a number\u2019s square root.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Str"),(0,u.kt)("td",null,"Returns a string representation of a number.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"StrComp"),(0,u.kt)("td",null,"Returns a value indicating the result of a string comparison.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"String"),(0,u.kt)("td",null,"Returns a repeating character or string.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Tan"),(0,u.kt)("td",null,"Returns a number\u2019s tangent.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Time"),(0,u.kt)("td",null,"Returns the current system time.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Timer"),(0,u.kt)("td",null,"Returns the number of seconds since midnight.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"TimeSerial"),(0,u.kt)("td",null,"Returns the time for a specified hour, minute, and second.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"TimeValue"),(0,u.kt)("td",null,"Converts a string to a time serial number.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Trim"),(0,u.kt)("td",null,"Returns a string without leading or trailing spaces.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"TypeName"),(0,u.kt)("td",null,"Returns a string that describes a variable\u2019s data type.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"UBound"),(0,u.kt)("td",null,"Returns the largest available subscript for an array\u2019s dimension.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"UCase"),(0,u.kt)("td",null,"Converts a string to uppercase.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Val"),(0,u.kt)("td",null,"Returns the numbers contained in a string.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"VarType"),(0,u.kt)("td",null,"Returns a value indicating a variable\u2019s subtype.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Weekday"),(0,u.kt)("td",null,"Returns a number representing a day of the week.")),(0,u.kt)("tr",null,(0,u.kt)("td",null,"Year"),(0,u.kt)("td",null,"Returns the year from a date value."))))}f.isMDXComponent=!0}}]);