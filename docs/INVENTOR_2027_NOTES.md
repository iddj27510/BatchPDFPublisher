# INVENTOR_2027_NOTES.md - VBA-DR-01 Batch PDF Publisher

**Version:** 9.25.42
**Last updated:** 2026-10-04

Notes on Inventor 2027 API compatibility, add-in manifest requirements,
and integration considerations.

---

## 1. Inventor 2027 API Novelties

Inventor 2027 introduces several API features not used by this
project (they are out of scope) but relevant for future planning:

| Feature | Purpose | Relevance |
|---|---|---|
| SlotFeature | Create slot features programmatically | N/A — project is read-only |
| ReplaceBOMView (Balloon, PartsList) | Replace BOM views programmatically | N/A |
| IrregularDistributionSet | Non-uniform distributions in patterns | N/A |
| GenerateOnSaveStatus | Control generation-on-save behavior | **Potential optimization for PDF_PrepareDocument** |
| BrowserFolder in Part docs | Organize browser nodes | N/A |

### GenerateOnSaveStatus — future optimization

`PDF_PrepareDocument` currently calls `oDoc.Save2 False` only if
`oDoc.Dirty`. If `GenerateOnSaveStatus` provides finer control, a
future version could skip regeneration of views that don't need
updating, reducing save time on large assemblies.

**Action:** Evaluate in a future minor release. Not required for
v9.25.42.

---

## 2. Add-in Manifest - UseInventorAssemblyContext (§9.6)

For deployment as compiled `.ivb` (Option B in README), the add-in
manifest must include the `UseInventorAssemblyContext` element. This
isolates the add-in's dependency context from Inventor's assembly
context, preventing conflicts with the Fusion Connected Client
architecture introduced in Inventor 2027.

### What it does

Without `UseInventorAssemblyContext`, the add-in loads its
dependencies into Inventor's main AppDomain. If Inventor also loads
a conflicting version of the same assembly (e.g., for Fusion), the
add-in may crash or misbehave.

With `UseInventorAssemblyContext=true`, the add-in gets its own
isolated assembly context. This is the recommended setting for all
Inventor 2027 add-ins.

### Manifest template

`docs/legacy/BatchPDFPublisher.addin.template` is the v8.0 template.
For v9.25.42 compiled deployment, use:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Addin Type="Standard">
  <Assembly>BatchPDFPublisher.dll</Assembly>
  <LoadAutomatically>0</LoadAutomatically>
  <UserUnloadable>1</UserUnloadable>
  <Hidden>0</Hidden>
  <SupportedSoftwareVersionGreaterThan>26..</SupportedSoftwareVersionGreaterThan>
  <DataVersion>1</DataVersion>
  <UseInventorAssemblyContext>true</UseInventorAssemblyContext>
  <ClassId>{B8A2F4D1-5E3C-4B7A-9F1E-2A4D6E8F0C12}</ClassId>
  <ClientId>{B8A2F4D1-5E3C-4B7A-9F1E-2A4D6E8F0C12}</ClientId>
  <DisplayName>Batch PDF Publisher</DisplayName>
  <Description>Bulk IDW to PDF export with drawing audit</Description>
  <Icon>BatchPDFPublisher.ico</Icon>
</Addin>

3. Autodesk Assistant (Inventor 2027)
Inventor 2027 ships with an embedded AI assistant ("Autodesk
Assistant"). It has no impact on this project — the add-in does not
expose iLogic rules, does not register a chat surface, and does not
participate in the assistant's tool registry.

If a future version integrates with the Assistant (e.g., to expose
"batch export" as a natural-language command), the following would
apply:

EU AI Act Article 50 disclosure obligations (transparency for
AI systems that interact with users) become relevant

OWASP LLM Top 10 2026 checklist applies to the integration layer

modAI was removed in v9.16.0 as dead code; re-introducing AI
would require a documented design decision

Status: Not planned. Documented here for future reference only.

4. .NET 8 Add-in Compatibility
Inventor 2027 supports .NET 8 add-ins without recompilation from
the .NET 6 baseline. This is relevant only if the project migrates
from VBA to .NET (see ROADMAP.md v10.x section). Current project
is pure VBA and does not use this capability.

5. API Timeouts
Inventor API calls are in-process COM. There is no configurable
timeout at the API level. Long-running operations (opening a large
IDW, updating a complex assembly) block the calling thread until
the operation completes or the OS-level COM timeout fires.

The project mitigates this by:

Throttled DoEvents via SC_TickProgress (folder scan)

Throttled DoEvents via PDF_ThrottledDoEvents (PDF export)

Cancel flag g_bCancel polled at file boundaries

There is no way to interrupt a single Inventor API call once it has
started. See KNOWN_LIMITATIONS.md "Cancel at File Boundary".

End of INVENTOR_2027_NOTES.md