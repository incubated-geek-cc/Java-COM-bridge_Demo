<div align="center">
  <img src='https://raw.githubusercontent.com/incubated-geek-cc/Java-COM-bridge_Demo/main/img/logo.png' width='96' height='96' alt='logo' />
  <h1 dir="auto">Java-COM bridge (Jacob)</h1>

**📀 Enables Java to interface with Microsoft Office applications for task automation.**

<div align="left">

<img src='https://raw.githubusercontent.com/incubated-geek-cc/Java-COM-bridge_Demo/main/img/jacob_diagram.png' width="800px" />

### 📌 Features (WIP)

</div>
<div align="left">
<ul>
	<li>📧 Email parsing on Outlook Exchange client.</li>
</ul>

#### For a full list of functionalities for Outlook, refer to <a href="https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.application" target="_blank">Office Outlook Primary Interop Assemply</a>. (namespaces with properties & methods have been compiled at <a href="https://raw.githubusercontent.com/incubated-geek-cc/Java-COM-bridge_Demo/main/office_outlook_interop_assembly.json" target="_blank">office_outlook_interop_assembly.json</a>.)

</div>

</div>

### 🌟 Try it yourself

<div align="left">
<ol>
	<li>Run <code>compile.sh</code> to generate <code>Main</code> class and then proceed to run <code>run.sh</code></li>
	<li><code>Main</code> class shall retrieve all inbox emails with subject title containing <code>Case Notification</code> and <code>CaseID</code> and output all JSON-formatted case details onto console.</li>
</ol>

#### Assume that embedded case details follow the below format:
<img src='https://raw.githubusercontent.com/incubated-geek-cc/Java-COM-bridge_Demo/main/img/sample_case_details.png' width="800px" />

#### Dependencies:
<table>
	<thead>
		<tr><th>Library</th><th>Description</th></tr>
	</thead>
	<tbody>
		<tr><td><code>jacob-1.18-x64.dll</code> and <code>lib/jacob-1.18.jar</code></td><td>For interfacing with Microsoft applications.</td></tr>
		<tr><td><code>lib/jsoup-1.15.3.jar</code> </td><td>For parsing HTML-formatted email message.</td></tr>
		<tr><td><code>lib/json-org-20140107.jar</code></td><td>To format processed output into <code>JSON</code></td></tr>
	</tbody>
</table>
</div>

### ✍ Read related posts here
[**Article :: Link :: Automate Outlook Email Tasks Using Jacob (Java-COM bridge)**](https://mobileappcircular.com/automate-outlook-email-tasks-using-jacob-java-com-bridge-3cf84ced2286)
<br>
<img src='https://raw.githubusercontent.com/incubated-geek-cc/Java-COM-bridge_Demo/main/img/overall_workflow.png' width="600px" />
<br><br>
[**Article :: Link :: Sending and Reading Outlook Emails Using Java - Jacob (Java-COM bridge)**](https://geek-cc.medium.com/sending-and-reading-outlook-emails-using-java-jacob-java-com-bridge-87f400bb2afc)
<br>
<img src='https://raw.githubusercontent.com/incubated-geek-cc/Java-COM-bridge_Demo/main/img/part_ii_use_cases.png' width="600px" />

<p>— <b>Join me on 📝 <b>Medium</b> at <a href='https://medium.com/@geek-cc' target='_blank'>~ ξ(🎀˶❛◡❛) @geek-cc</a></b></p>

---

#### 🌮 Please buy me a <a href='https://www.buymeacoffee.com/geekcc' target='_blank'>Taco</a>! 😋


## 📜 License & Credits

<ol>
	<li>Original library at <a href="https://github.com/freemansoft/jacob-project" target="_blank">jacob-project</a> by 👤 <a href="https://github.com/freemansoft" target="_blank">Joe Freeman (freemansoft)</a></li>
</ol>