# 🔍 DIAGNOSTIC QUESTIONS FOR O3: Excel COM Add-in Runtime Error

## PROBLEM STATEMENT
Excel COM add-in "Snipper Pro" shows "Not loaded. A runtime error occurred during the loading of the COM Add-in" with location "mscoree.dll". The add-in appears in COM Add-ins list but fails to load, setting LoadBehavior to 2 (disabled due to error).

## CURRENT STATE
- ✅ COM instantiation works: `New-Object -ComObject "SnipperPro.Connect"` succeeds
- ✅ Registry entries correct: LoadBehavior=3, proper CLSID/ProgID registered
- ✅ DLL exists and is accessible
- ❌ Excel cannot load the add-in (runtime error in mscoree.dll)

---

## 1. INTERFACE IMPLEMENTATION QUESTIONS

**Q1.1:** Does our COM add-in class implement the correct interface for Excel COM add-ins?
- Current: We implement `IRibbonExtensibility` only
- Question: Should we implement `IDTExtensibility2` for proper COM add-in lifecycle?
- Missing methods: `OnConnection`, `OnDisconnection`, `OnAddInsUpdate`, `OnStartupComplete`, `OnBeginShutdown`

**Q1.2:** Are we missing required COM interface signatures?
- Check if our `OnConnection` signature matches Excel's expectations
- Verify parameter types: `object application, int connectMode, object addInInst, ref Array custom`

**Q1.3:** Should our class inherit from a specific base class or implement additional interfaces?

---

## 2. COM REGISTRATION QUESTIONS

**Q2.1:** Is our COM registration complete for Excel integration?
- Do we need additional registry entries beyond CLSID and ProgID?
- Are we missing `TypeLib` registration?
- Do we need `InprocServer32` entries?

**Q2.2:** What specific registry structure does Excel expect for COM add-ins?
- Compare our registry with working Excel COM add-ins
- Check if we need specific threading model settings

**Q2.3:** Are there Excel-specific COM registration requirements we're missing?
- Office Trust Center settings
- COM security settings
- 32-bit vs 64-bit registration issues

---

## 3. DEPENDENCY AND ASSEMBLY QUESTIONS

**Q3.1:** What are the minimal required assemblies for an Excel COM add-in?
- Current refs: Office.dll, Microsoft.Office.Interop.Excel.dll, System, System.Core, System.Windows.Forms
- Missing: Do we need Microsoft.Office.Tools.Common?

**Q3.2:** Are there .NET Framework version compatibility issues?
- Current: .NET Framework 4.8
- Question: Does Excel version require specific .NET Framework version?

**Q3.3:** Should assemblies be GAC registered or use specific binding redirects?

**Q3.4:** Are we missing required Office Primary Interop Assemblies (PIAs)?

---

## 4. ASSEMBLY ATTRIBUTES AND METADATA QUESTIONS

**Q4.1:** What COM attributes are required for Excel COM add-ins?
- Current: `[ComVisible(true)]`, `[Guid]`, `[ProgId]`, `[ClassInterface(ClassInterfaceType.None)]`
- Missing: Do we need `[ComDefaultInterface]`, `[ComSourceInterfaces]`?

**Q4.2:** Should we specify explicit interface definitions?
- Create explicit COM interface contracts?
- Use `[InterfaceType(ComInterfaceType.InterfaceIsDual)]`?

**Q4.3:** Are there required AssemblyInfo attributes for COM interop?

---

## 5. INITIALIZATION AND LIFECYCLE QUESTIONS

**Q5.1:** What is the exact initialization sequence Excel expects?
- When should ribbon XML be returned?
- What should happen in `OnConnection` vs `OnRibbonLoad`?

**Q5.2:** Are we handling Excel's threading model correctly?
- STA vs MTA apartment state issues
- Thread safety of our implementation

**Q5.3:** Should we defer heavy initialization until after `OnConnection` completes?

---

## 6. DEBUGGING AND ERROR HANDLING QUESTIONS

**Q6.1:** How can we capture the actual runtime exception occurring in mscoree.dll?
- Enable fusion logging for assembly binding failures
- Use COM error handling to get detailed error info
- Add logging to every method entry/exit

**Q6.2:** What debugging techniques work for COM add-in loading failures?
- Attach debugger during Excel startup
- Use ProcMon to monitor file/registry access
- Enable .NET Framework debugging

**Q6.3:** Are there Excel event logs or COM-specific logs we should check?

---

## 7. MINIMAL WORKING EXAMPLE QUESTIONS

**Q7.1:** What is the absolute minimal COM add-in that Excel can load successfully?
- Simplest possible class implementation
- Minimal registry entries
- Basic ribbon returning empty XML

**Q7.2:** Can we start with a stub implementation and gradually add functionality?

**Q7.3:** What COM add-in templates or examples are known to work with current Excel versions?

---

## 8. EXCEL VERSION AND ENVIRONMENT QUESTIONS

**Q8.1:** Are there Excel version-specific requirements for COM add-ins?
- Office 365 vs standalone Office versions
- 32-bit vs 64-bit Office compatibility

**Q8.2:** Do we need to handle different Office security models?
- Trust Center settings
- VBA macro security affecting COM add-ins
- Digital signing requirements

**Q8.3:** Are there Windows version or .NET security policies affecting COM loading?

---

## 9. ALTERNATIVE APPROACH QUESTIONS

**Q9.1:** Should we try VSTO approach instead of pure COM?
- Create proper VSTO manifest
- Use ClickOnce deployment
- Leverage VSTO runtime

**Q9.2:** Can we use Office Add-in model (Office.js) instead?
- Web-based add-in approach
- Modern Office Add-ins platform

**Q9.3:** Should we implement as Excel-DNA add-in?
- .NET-based Excel add-in framework
- Might handle COM complexities automatically

---

## 10. IMMEDIATE DIAGNOSTIC ACTIONS

**Q10.1:** Can you provide the exact steps to:
1. Enable fusion logging to capture assembly binding failures
2. Use Process Monitor to see what files/registry Excel accesses during add-in loading
3. Attach Visual Studio debugger to Excel process during add-in loading
4. Check Windows Event Viewer for detailed COM errors

**Q10.2:** What specific registry comparison should we do?
- Export registry from machine with working Excel COM add-in
- Compare with our current registration

**Q10.3:** Should we test with a completely different GUID to ensure no cached registration issues?

---

## PRIORITY ORDER FOR O3:
1. **Interface Implementation** - Most likely cause
2. **COM Registration completeness** - Second most likely
3. **Assembly dependencies** - Third most likely
4. **Debugging to capture actual error** - Essential for diagnosis
5. **Minimal working example** - Fallback approach

## REQUEST FOR O3:
Please systematically address these questions starting with the highest priority items. Provide specific code examples, registry entries, and step-by-step debugging procedures to identify and resolve the mscoree.dll runtime error in Excel COM add-in loading. 