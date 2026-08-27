// Builds the "Site RPs" sheet for a tracker workbook entirely from shared sources, so adding
// a store needs no manual worksheet editing in either tracker.
//
// Paste into Power Query's Advanced Editor and name the query "Site RPs".
//
// REPLACES what used to be a hand-typed store list plus XLOOKUP formulas pointing at
// [Store AP Staff Directory.xlsx]. Those links drifted: JCI's resolved to the SharePoint copy
// while Securitas's was hardcoded to C:\Users\p4bn\Downloads\, leaving the two trackers
// disagreeing about the responsible party for 100 stores. A query against shared sources
// cannot drift.
//
// Two sources, both already on SharePoint:
//   BU List                        the store list, store name and BU (1079 unique stores)
//   Store AP Staff Directory.xlsx  tier 1/2/3 contacts (439 stores)
//
// Emits the same 11 columns in the same order as before, so TenantSheetName("site-rps") and
// the VBA cascade in RequesterEmail.vb need no change.
let
    // ---- helpers -----------------------------------------------------------------
    // Everything is emitted as text. A null becomes "" so the VBA sees a blank cell rather
    // than the string "null".
    AsText = (v) => if v = null then "" else Text.Trim(Text.From(v)),

    // The old formulas treated "N/A" and "OPEN" as "no contact here", alongside blank.
    // Preserved exactly -- these are real values in the directory, not data errors.
    Usable = (v) =>
        let t = Text.Upper(AsText(v))
        in t <> "" and t <> "N/A" and t <> "OPEN",

    // First usable of the three, falling back to the THIRD even when it is unusable. That is
    // what the old nested IF did, so "N/A" could legitimately surface in Responsible/Email;
    // reproduced rather than "improved" so the swap changes nothing it does not have to.
    Pick = (a, b, c) => if Usable(a) then AsText(a) else if Usable(b) then AsText(b) else AsText(c),

    Key4 = (v) => Text.PadStart(AsText(v), 4, "0"),

    // ---- source A: the store list, from the BU List query already in this workbook ----
    // Referencing the existing query rather than re-querying SharePoint keeps one definition
    // of "what stores exist" per workbook.
    Stores = Table.SelectColumns(#"BU List", {"Store", "STORE NAME", "BU"}),
    StoresKeyed = Table.AddColumn(Stores, "StoreKey", each Key4([Store]), type text),

    // ---- source B: the AP Staff Directory on the same SharePoint site ----------------
    SiteFiles = SharePoint.Files(
        "https://nordstrom.sharepoint.com/sites/AssetProtectionPMO",
        [ApiVersion = 15]
    ),
    DirBinary = Table.SelectRows(
        SiteFiles,
        each [Name] = "Store AP Staff Directory.xlsx"
    ){0}[Content],
    // Read the SHEET rather than its table, so renaming the table cannot break this.
    DirSheet = Excel.Workbook(DirBinary, null, true){[Item = "AP POCs", Kind = "Sheet"]}[Data],
    DirHeaders = Table.PromoteHeaders(DirSheet, [PromoteAllScalars = true]),
    DirCols = Table.SelectColumns(
        DirHeaders,
        {"Store", "Store Name", "Tier 1 Name", "Tier 1 Email",
         "Tier 2 Name", "Tier 2 Email", "Tier 3 Name", "Tier 3 Email"}
    ),
    DirKeyed = Table.AddColumn(DirCols, "StoreKey", each Key4([Store]), type text),
    Dir = Table.RemoveColumns(DirKeyed, {"Store"}),

    // ---- join: every store in BU List, with contacts where the directory has them ----
    Joined = Table.NestedJoin(StoresKeyed, {"StoreKey"}, Dir, {"StoreKey"}, "dir", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(
        Joined, "dir",
        {"Store Name", "Tier 1 Name", "Tier 1 Email", "Tier 2 Name", "Tier 2 Email",
         "Tier 3 Name", "Tier 3 Email"},
        {"DirName", "D1N", "D1E", "D2N", "D2E", "D3N", "D3E"}
    ),

    // ---- Vertical, derived from BU --------------------------------------------------
    // Was hand-typed. Only NS-ness is ever tested downstream, so NR / SC / Corporate / blank
    // all behave identically -- which is why deriving this changed the effective vertical for
    // 223 stores but the actual outcome for only 5, all of which were accepted deliberately:
    // 0391, 0395, 4005 become NS; 5501, 5502 become NR.
    WithVertical = Table.AddColumn(Expanded, "VerticalCalc", each
        if AsText([BU]) = "200" then "NS"
        else if AsText([BU]) = "250" then "NR"
        else "", type text),

    // ---- tier columns ---------------------------------------------------------------
    // Tier 1 is suppressed for NS stores. For a full-line store the directory's tier 1 is the
    // store-level AP person, who is not the right requisition requester -- tier 2 is. This was
    // the =IF(E2="NS","",XLOOKUP(...)) in the old sheet and it is load-bearing: without it,
    // NS stores would start raising requisitions for a different person.
    WithTiers = Table.AddColumn(WithVertical, "T1NameCalc",
                    each if [VerticalCalc] = "NS" then "" else AsText([D1N]), type text),
    WithTiers2 = Table.AddColumn(WithTiers, "T1EmailCalc",
                    each if [VerticalCalc] = "NS" then "" else AsText([D1E]), type text),

    // ---- the derived Responsible / Email --------------------------------------------
    WithResp = Table.AddColumn(WithTiers2, "ResponsibleCalc",
                   each Pick([T1NameCalc], [D2N], [D3N]), type text),
    WithEmail = Table.AddColumn(WithResp, "EmailCalc",
                    each Pick([T1EmailCalc], [D2E], [D3E]), type text),

    // ---- shape it exactly like the sheet it replaces --------------------------------
    // Store Name prefers the directory's casing ("Downtown Seattle") over BU List's
    // ("DOWNTOWN SEATTLE"), which is what the hand-maintained sheet showed.
    WithName = Table.AddColumn(WithEmail, "StoreNameCalc",
                   each if Usable([DirName]) then AsText([DirName]) else AsText([#"STORE NAME"]),
                   type text),

    Shaped = Table.SelectColumns(WithName,
        {"StoreKey", "StoreNameCalc", "ResponsibleCalc", "EmailCalc", "VerticalCalc",
         "T1NameCalc", "T1EmailCalc", "D2N", "D2E", "D3N", "D3E"}),

    Renamed = Table.RenameColumns(Shaped, {
        {"StoreKey", "Store"},
        {"StoreNameCalc", "Store Name"},
        {"ResponsibleCalc", "Responsible"},
        {"EmailCalc", "Email"},
        {"VerticalCalc", "Vertical"},
        {"T1NameCalc", "Tier 1 name"},
        {"T1EmailCalc", "Tier 1 email"},
        {"D2N", "Tier 2 name"},
        {"D2E", "Tier 2 email"},
        {"D3N", "Tier 3 name"},
        {"D3E", "Tier 3 email"}
    }),

    // Nulls from the left join become "", so no cell reads "null" to the VBA or to a human.
    Filled = Table.TransformColumns(Renamed,
        List.Transform(Table.ColumnNames(Renamed), each {_, AsText, type text})),

    // Sorted so the sheet is stable between refreshes and diffable.
    Sorted = Table.Sort(Filled, {{"Store", Order.Ascending}})
in
    Sorted
