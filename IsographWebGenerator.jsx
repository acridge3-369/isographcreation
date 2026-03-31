import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Download, FileSpreadsheet, Loader2, Upload } from "lucide-react";

const PAGE_NAME = "SYSTEM";
const DEFAULT_LABOR = "Marine technician";
const DEFAULT_ACTIVE_TIME = 0.5;

const RBD_BLOCK_X_STEP = 160;
const RBD_BLOCK_X_MIN = 0;
const RBD_BLOCK_X_MAX = 640;
const RBD_BLOCK_Y_START = 100;
const RBD_BLOCK_Y_STEP = 100;
const RBD_BLOCK_WIDTH = 100;
const BLOCKS_PER_ROW = Math.floor((RBD_BLOCK_X_MAX - RBD_BLOCK_X_MIN) / RBD_BLOCK_X_STEP) + 1;

function makeGuid() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

function cleanText(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function cleanNumber(value, defaultValue = 0) {
  if (value === null || value === undefined || value === "") return defaultValue;
  const n = Number(value);
  return Number.isFinite(n) ? n : defaultValue;
}

function normalizeTradeName(value) {
  const text = cleanText(value);
  if (!text || ["N/A", "OEM"].includes(text.toUpperCase())) return DEFAULT_LABOR;
  return text;
}

function classifyItemRef(value) {
  const text = cleanText(value);
  if (/^\d+$/.test(text)) return "main";
  if (/^\d+\.\d+$/.test(text)) return "child";
  return "other";
}

function normalizeHeader(value) {
  return cleanText(value).toLowerCase().replace(/\s+/g, " ").trim();
}

function findColumnKey(row, candidates) {
  const keys = Object.keys(row || {});
  const normalizedKeys = keys.map((key) => ({ raw: key, norm: normalizeHeader(key) }));

  for (const candidate of candidates) {
    const wanted = normalizeHeader(candidate);
    const exact = normalizedKeys.find((item) => item.norm === wanted);
    if (exact) return exact.raw;
  }

  for (const candidate of candidates) {
    const wanted = normalizeHeader(candidate);
    const partial = normalizedKeys.find((item) => item.norm.includes(wanted) || wanted.includes(item.norm));
    if (partial) return partial.raw;
  }

  return null;
}

function getValue(row, candidates, fallback = "") {
  const key = findColumnKey(row, candidates);
  if (!key) return fallback;
  return row[key] ?? fallback;
}

function rowHasUsefulData(row) {
  return Object.values(row || {}).some((value) => cleanText(value) !== "");
}

function getSheetHeaders(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const headers = {};
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const cellAddress = XLSX.utils.encode_cell({ r: 0, c });
    const cell = sheet[cellAddress];
    if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== "") {
      headers[String(cell.v).trim()] = c;
    }
  }
  return headers;
}

function clearSheetData(sheet, startRow = 1) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  for (let r = startRow; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = XLSX.utils.encode_cell({ r, c });
      delete sheet[addr];
    }
  }
  sheet["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: range.s.c },
    e: { r: Math.max(startRow, range.e.r), c: range.e.c },
  });
}

function writeRecords(sheet, records, startRow = 1) {
  const headers = getSheetHeaders(sheet);
  let maxRow = startRow;
  let maxCol = 0;

  Object.values(headers).forEach((c) => {
    if (c > maxCol) maxCol = c;
  });

  records.forEach((record, index) => {
    const row = startRow + index;
    Object.entries(record).forEach(([key, value]) => {
      const col = headers[key];
      if (col === undefined) return;
      const addr = XLSX.utils.encode_cell({ r: row, c: col });
      sheet[addr] = {
        t: typeof value === "number" ? "n" : typeof value === "boolean" ? "b" : "s",
        v: value,
      };
      if (col > maxCol) maxCol = col;
    });
    if (row > maxRow) maxRow = row;
  });

  sheet["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } });
}

function getRbdBlockPosition(blockIndex) {
  const row = Math.floor(blockIndex / BLOCKS_PER_ROW);
  const colInRow = blockIndex % BLOCKS_PER_ROW;
  const isEvenRow = row % 2 === 0;
  const x = isEvenRow
    ? RBD_BLOCK_X_MIN + colInRow * RBD_BLOCK_X_STEP
    : RBD_BLOCK_X_MAX - colInRow * RBD_BLOCK_X_STEP;
  const y = RBD_BLOCK_Y_START + row * RBD_BLOCK_Y_STEP;
  return { x, y };
}

function renameFMECAColumns(rows) {
  return rows.map((row, index) => ({
    ...row,
    __RowIndex: index + 1,
    ItemRef: getValue(row, ["Item Ref. No.", "Item Ref", "Item", "Ref", "Reference", "ID"], ""),
    Equipment: getValue(row, ["EQUIPMENT", "Equipment"], ""),
    AbsLocation: getValue(row, ["Abs Location", "Location"], ""),
    CommonName: getValue(row, ["CCG 5-Point Naming Standard", "Common Name", "Equipment Name", "Name", "Title", "Tag"], ""),
    Brand: getValue(row, ["Brand"], ""),
    CombinedNamingStandard: getValue(row, ["SYSTEM", "System", "Combined Naming Standard", "Combined Name", "System Name"], ""),
    SupplierName: getValue(row, ["Supplier (SSI or SME)", "Supplier", "Manufacturer"], ""),
    SupplierEquipmentName: getValue(row, ["Supplier Equipment Name", "Supplier Name"], ""),
    SupplierPartNumber: getValue(row, ["Supplier Part Number", "Part Number", "Part No", "PN"], ""),
    VendorName: getValue(row, ["OEM / Vendor", "Vendor", "OEM"], ""),
    VendorEquipmentName: getValue(row, ["Vendor Equipment Name", "OEM Name"], ""),
    VendorModelNumber: getValue(row, ["Vendor Model Number", "Model Number", "Model"], ""),
    VendorPartNumber: getValue(row, ["Vendor Part Number", "Vendor PN", "OEM Part Number"], ""),
    EquipmentType: getValue(row, ["Equipment Identification (COTS, non-COTS, etc)", "Equipment Type", "Type"], ""),
    TaskName: getValue(row, ["FAILURE MODE", "Failure Mode", "Task Name", "Task", "Maintenance Task"], ""),
    TaskDescription: getValue(row, ["FAILURE EFFECT", "Failure Effect", "Task Description", "Description", "Failure Description", "Details"], ""),
    TaskTitle: getValue(row, ["Task Title", "Title"], ""),
    TaskType: getValue(row, ["Task Type", "Type", "Task Classification"], ""),
    TaskFrequency: getValue(row, ["Task Frequency", "Frequency"], ""),
    TaskSource: getValue(row, ["Task Source", "Source"], ""),
    TaskRationale: getValue(row, ["Task Rationale", "Rationale"], ""),
    TaskInterval: getValue(row, ["Task Interval", "Interval"], ""),
    JobStepTitle: getValue(row, ["Job Step Title", "Step Title"], ""),
    MaintenanceLine: getValue(row, ["Maintenance Line", "Maintenance"], ""),
    Procedure: getValue(row, ["Procedure", "Instructions", "Method"], ""),
    RequiredParts: getValue(row, ["PART NUMBER", "Required Parts", "Req parts, materials, tools, test equipment", "Parts", "Materials", "Spares"], ""),
    SafetyPrecautions: getValue(row, ["Safety", "Safety Precautions"], ""),
    EnvironmentalRequirements: getValue(row, ["Environmental Requirements"], ""),
    PreMaintenanceConditions: getValue(row, ["Pre Maintenance Conditions", "Pre-Maintenance Conditions"], ""),
    RegulationsAndStandards: getValue(row, ["Regulations And Standards", "Standards"], ""),
    TechnicalData: getValue(row, ["Technical Data"], ""),
    TradeCategory: getValue(row, ["Required Resources", "Trade Category"], ""),
    TradeName: getValue(row, ["Trade Name", "Labor", "Labour", "Required Resources", "Technician", "Trade"], ""),
    TaskClassification: getValue(row, ["Task Classification", "Classification"], ""),
    LabourHours: getValue(row, ["Labour Hours", "Labor Hours", "Hours", "Active Time", "Task Duration"], ""),
    Comments: getValue(row, ["FMEA RECOMMENDED ACTION", "Comments", "Notes", "Remarks"], ""),
    Troubleshooting: getValue(row, ["Troubleshooting", "Trouble Shooting"], ""),
  }));
}

function buildRecords(fmecaRows) {
  const failureModels = [];
  const correctiveTasks = [];
  const correctiveTaskLabor = [];
  const correctiveTaskSpares = [];
  const scheduledTasks = [];
  const scheduledTaskLabor = [];
  const laborRows = [];
  const spareRows = [];
  const rbdBlocks = [];
  const rbdConnections = [];
  const rbdNodes = [];
  const taskGroups = [];

  const seenLabor = new Set();
  const seenSpares = new Set();

  let currentFailureModel = null;
  let currentIndex = 0;
  const blockIds = [];

  for (const row of fmecaRows) {
    if (!rowHasUsefulData(row)) continue;

    const rowType = classifyItemRef(row.ItemRef);
    if (rowType !== "main" && rowType !== "child") continue;

    if (rowType === "main") {
      const equipmentId = cleanText(row.Equipment);
      if (!equipmentId) continue;

      currentIndex += 1;
      const systemId = `${PAGE_NAME}.${currentIndex}`;
      currentFailureModel = systemId;
      blockIds.push(systemId);

      const commonName = cleanText(row.CommonName);
      const vendorName = cleanText(row.VendorName);
      const vendorPart = cleanText(row.VendorPartNumber);
      const taskName = cleanText(row.TaskName);
      const taskDesc = cleanText(row.TaskDescription);
      const taskTitle = cleanText(row.TaskTitle);
      const taskType = cleanText(row.TaskType);
      const taskFreq = cleanText(row.TaskFrequency);
      const taskInterval = cleanText(row.TaskInterval);
      const taskSource = cleanText(row.TaskSource);
      const taskRationale = cleanText(row.TaskRationale);
      const procedure = cleanText(row.Procedure);
      const troubleshooting = cleanText(row.Troubleshooting);
      const comments = cleanText(row.Comments);
      const requiredParts = cleanText(row.RequiredParts);
      const labourHours = cleanNumber(row.LabourHours, DEFAULT_ACTIVE_TIME);
      const tradeName = normalizeTradeName(row.TradeName);

      const description = taskName || equipmentId || systemId;
      const fullDescription = taskDesc || procedure || description;
      const externalReference = [commonName, vendorName, vendorPart].filter(Boolean).join(" | ");

      // FailureModels.Id uses Equipment from FMECA.
      failureModels.push({
        Id: equipmentId,
        AlCapitalCost: 0,
        AlCostRate: 0,
        AlDescription: description,
        AlDetectionP: 1,
        AlEnabled: false,
        AlPfDistribution: "Step",
        AlPfInterval: 0,
        AlPfStd: 0,
        CoCapitalCost: 0,
        CoCostRate: 0,
        CoDescription: fullDescription,
        CoEnabled: false,
        Description: description,
        ExternalReference: externalReference || systemId,
        FmDistribution: "Exponential",
        FmDormant: false,
        FmEta1: 0,
        FmGamma1: 0,
        FmMttf: 0,
        Guid: makeGuid(),
        Notes1: taskTitle,
        Notes2: taskSource,
        Notes3: taskRationale,
        Notes4: troubleshooting || comments,
        Remarks: requiredParts,
        StandbyAgeingPercent: 100,
        StandbyFailurePercent: 100,
        StartUpFailureProb: 0,
        Type: "Failure model",
      });

      // All related sheets use system.n sequence in Id columns.
      correctiveTasks.push({
        Id: systemId,
        AgeReductionFactor: 0,
        AgeReductionMode: "As good as new",
        Beta: 2,
        Description: description,
        Distribution: "Normal",
        EquipmentDescriptions: equipmentId,
        EquipmentIds: systemId,
        Eta: Math.max(labourHours, DEFAULT_ACTIVE_TIME),
        ExternalReference: externalReference,
        FailureModel: systemId,
        Gamma: 0,
        LaborDescriptions: tradeName,
        LaborIds: tradeName,
        OperationalCost: 0,
        OperationNumber: 1,
        RampTime: 0,
        SparesDescriptions: requiredParts,
        SparesIds: "",
        Std: 0,
        TaskDuration: Math.max(labourHours, DEFAULT_ACTIVE_TIME),
        TaskId: systemId,
        WeibullSet: false,
      });

      correctiveTaskLabor.push({
        ActiveTime: Math.max(labourHours, DEFAULT_ACTIVE_TIME),
        ExternalReference: externalReference,
        FailureModel: systemId,
        Labor: tradeName,
        Quantity: 1,
        SetToTaskDuration: true,
        SubIndex: 0,
      });

      scheduledTasks.push({
        Id: systemId,
        AgeReductionFactor: 0,
        AgeReductionMode: "As good as new",
        Baseline: false,
        Beta: 2,
        Description: description,
        DetectionP: 1,
        Distribution: "Normal",
        Enabled: true,
        EquipmentDescriptions: equipmentId,
        EquipmentIds: systemId,
        Eta: 24,
        ExternalReference: externalReference,
        FailureModel: systemId,
        FixedInterval: false,
        Gamma: 0,
        LaborDescriptions: tradeName,
        LaborIds: tradeName,
        Mandatory: false,
        MinimumAge: 0,
        Notes1: taskTitle,
        Notes2: taskSource,
        Notes3: taskRationale,
        Notes4: procedure,
        Offset: 0,
        OperationalCost: 0,
        OperationNumber: 1,
        OutedDuringMaintenance: false,
        PfDistribution: "Step",
        PfInterval: 0,
        PfStd: 0,
        RampTime: 0,
        SparesDescriptions: requiredParts,
        SparesIds: "",
        Std: 0,
        SubIndex: 0,
        TaskDuration: Math.max(labourHours, DEFAULT_ACTIVE_TIME),
        TaskGroup: "",
        TaskId: systemId,
        TaskInterval: taskInterval || taskFreq,
        TaskTrigger: "Age",
        Type: taskType || "PM",
        WeibullSet: false,
      });

      scheduledTaskLabor.push({
        ActiveTime: Math.max(labourHours, DEFAULT_ACTIVE_TIME),
        ExternalReference: externalReference,
        FailureModel: systemId,
        Labor: tradeName,
        Quantity: 1,
        SecondarySubIndex: 0,
        SetToTaskDuration: true,
        SubIndex: 0,
      });

      if (!seenLabor.has(tradeName)) {
        seenLabor.add(tradeName);
        laborRows.push({
          Id: tradeName,
          CostRate: 0,
          Description: tradeName,
          ExternalReference: "",
          Guid: makeGuid(),
          NoAvailable: 1,
          ScheduledCalloutCost: 0,
          CorrectiveCalloutCost: 0,
          CorrectiveLogisticDelay: 0,
          Type: "Labor",
        });
      }
    } else if (rowType === "child" && currentFailureModel) {
      const spareName = cleanText(row.CommonName);
      const supplierPart = cleanText(row.SupplierPartNumber);
      const vendorPart = cleanText(row.VendorPartNumber);
      const comments = cleanText(row.Comments);

      let spareId = supplierPart || vendorPart || spareName;
      const spareDesc = spareName || supplierPart || vendorPart || "Spare item";
      if (!spareId) spareId = `SPARE_${spareRows.length + 1}`;

      if (!seenSpares.has(spareId)) {
        seenSpares.add(spareId);
        spareRows.push({
          Id: spareId,
          Description: spareDesc,
          ExternalReference: comments,
          Guid: makeGuid(),
          RepairLevel: "Discard",
          SourceDistribution: "Constant",
          SourceStd: 0,
          Type: "Spare",
          UnitCost: 0,
          Volume: 0,
          Weight: 0,
        });
      }

      correctiveTaskSpares.push({
        ExternalReference: "",
        FailureModel: currentFailureModel,
        Quantity: 1,
        Spare: spareId,
        SubIndex: 0,
      });
    }
  }

  rbdBlocks.push({
    Id: PAGE_NAME,
    XPagePosition: 74,
    XPosition: 0,
  });

  blockIds.forEach((blockId, blockIndex) => {
    const { x, y } = getRbdBlockPosition(blockIndex);
    rbdBlocks.push({
      Id: blockId,
      BackgroundColor: -1,
      Description: blockId,
      ExternalReference: "",
      FailureModel: blockId,
      Guid: makeGuid(),
      Height: 60,
      IncludeInSystemPlot: true,
      Page: PAGE_NAME,
      PageScale: 100,
      Width: RBD_BLOCK_WIDTH,
      XPosition: x,
      YPosition: y,
    });
  });

  rbdNodes.push({
    Id: `${PAGE_NAME}.START`,
    BackgroundColor: -1,
    Guid: makeGuid(),
    LocalStandby: false,
    NotLogic: false,
    OperationalCapacityTarget: "",
    OperationalCapacityTarget2: "",
    OperationalCapacityTarget3: "",
    OperationalCapacityTarget4: "",
    Page: PAGE_NAME,
    Vote: 1,
    XPosition: 0,
    YPosition: 100,
  });
  rbdNodes.push({
    Id: `${PAGE_NAME}.END`,
    BackgroundColor: -1,
    Guid: makeGuid(),
    LocalStandby: false,
    NotLogic: false,
    OperationalCapacityTarget: "",
    OperationalCapacityTarget2: "",
    OperationalCapacityTarget3: "",
    OperationalCapacityTarget4: "",
    Page: PAGE_NAME,
    Vote: 1,
    XPosition: 100 + Math.max(blockIds.length - 1, 0) * 200 + 150,
    YPosition: 100,
  });

  let connectionId = 1;
  if (blockIds.length) {
    rbdConnections.push({
      Id: `${PAGE_NAME}.${connectionId}`,
      Color: -1,
      Guid: makeGuid(),
      InputObjectIndex: 0,
      InputObjectType: "RBD node",
      OutputObjectIndex: 0,
      OutputObjectType: "RBD block",
      Page: PAGE_NAME,
      Type: "Horizontal/vertical",
    });
    connectionId += 1;

    for (let i = 0; i < blockIds.length - 1; i += 1) {
      rbdConnections.push({
        Id: `${PAGE_NAME}.${connectionId}`,
        Color: -1,
        Guid: makeGuid(),
        InputObjectIndex: i,
        InputObjectType: "RBD block",
        OutputObjectIndex: i + 1,
        OutputObjectType: "RBD block",
        Page: PAGE_NAME,
        Type: "Horizontal/vertical",
      });
      connectionId += 1;
    }

    rbdConnections.push({
      Id: `${PAGE_NAME}.${connectionId}`,
      Color: -1,
      Guid: makeGuid(),
      InputObjectIndex: blockIds.length - 1,
      InputObjectType: "RBD block",
      OutputObjectIndex: 1,
      OutputObjectType: "RBD node",
      Page: PAGE_NAME,
      Type: "Horizontal/vertical",
    });
  }

  return {
    FailureModelCorrectiveTasks: correctiveTasks,
    FailureModelCorrTaskLabor: correctiveTaskLabor,
    FailureModelCorrTaskSpares: correctiveTaskSpares,
    FailureModels: failureModels,
    FailureModelSchedTaskLabor: scheduledTaskLabor,
    FailureModelScheduledTasks: scheduledTasks,
    Labor: laborRows,
    RbdBlocks: rbdBlocks,
    RbdConnections: rbdConnections,
    RbdNodes: rbdNodes,
    Spares: spareRows,
    TaskGroups: taskGroups,
  };
}

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellDates: false });
}

function readSourceRows(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  let headerRowIndex = 0;
  for (let i = 0; i < Math.min(matrix.length, 20); i += 1) {
    const row = (matrix[i] || []).map((cell) => normalizeHeader(cell));
    const hasEquipment = row.includes("equipment");
    const hasFailureMode = row.some((cell) => cell.includes("failure mode"));
    if (hasEquipment && hasFailureMode) {
      headerRowIndex = i;
      break;
    }
  }

  return XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    range: headerRowIndex,
  });
}

function makeWorkbookBlob(workbook) {
  const arrayBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  return new Blob([arrayBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export default function IsographWebGenerator() {
  const [templateFile, setTemplateFile] = useState(null);
  const [sourceFile, setSourceFile] = useState(null);
  const [status, setStatus] = useState("Upload template and FMECA files to generate an output workbook.");
  const [error, setError] = useState("");
  const [readyName, setReadyName] = useState("");
  const [downloadUrl, setDownloadUrl] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);

  const canGenerate = useMemo(() => !!templateFile && !!sourceFile && !isGenerating, [templateFile, sourceFile, isGenerating]);

  async function handleGenerate() {
    setError("");
    setReadyName("");
    setIsGenerating(true);

    if (downloadUrl) {
      URL.revokeObjectURL(downloadUrl);
      setDownloadUrl("");
    }

    try {
      if (!templateFile || !sourceFile) {
        setError("Please upload both files first.");
        return;
      }

      setStatus("Reading uploaded files...");
      const templateWb = await readWorkbook(templateFile);
      const sourceWb = await readWorkbook(sourceFile);

      const sourceSheetNames = sourceWb.SheetNames || Object.keys(sourceWb.Sheets || {});
      const sourceSheetName = sourceSheetNames[0];
      if (!sourceSheetName) {
        throw new Error("The FMECA workbook does not contain any sheets.");
      }

      setStatus(`Reading FMECA rows from '${sourceSheetName}'...`);
      let fmecaRows = readSourceRows(sourceWb.Sheets[sourceSheetName]);
      fmecaRows = renameFMECAColumns(fmecaRows).filter((row) => cleanText(row.Equipment) !== "");

      if (!fmecaRows.length) {
        throw new Error("No rows with Equipment values were found in the FMECA sheet.");
      }

      setStatus("Building Isograph sheet data...");
      const recordsBySheet = buildRecords(fmecaRows);

      Object.entries(recordsBySheet).forEach(([sheetName, records]) => {
        const sheet = templateWb.Sheets[sheetName];
        if (!sheet) return;
        clearSheetData(sheet, 2);
        writeRecords(sheet, records, 2);
      });

      const outputName = `${sourceFile.name.replace(/\.xlsx?$/i, "")}_generated_isograph.xlsx`;
      const blob = makeWorkbookBlob(templateWb);
      const url = URL.createObjectURL(blob);

      setDownloadUrl(url);
      setReadyName(outputName);
      setStatus("Workbook generated. Click download.");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Something went wrong while generating the workbook.");
      setStatus("Generation failed.");
    } finally {
      setIsGenerating(false);
    }
  }

  return (
    <div className="min-h-screen bg-slate-50 p-6">
      <div className="mx-auto grid max-w-5xl gap-6">
        <Card className="rounded-2xl shadow-sm">
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-2xl">
              <FileSpreadsheet className="h-6 w-6" />
              Isograph Template Generator
            </CardTitle>
          </CardHeader>
          <CardContent className="grid gap-6">
            <p className="text-sm text-slate-600">
              Step 1: Upload your master template workbook. Step 2: Upload your FMECA workbook. Step 3: Click Generate and then Download.
            </p>

            <div className="grid gap-4 md:grid-cols-2">
              <div className="grid gap-2">
                <Label htmlFor="template">Master template workbook</Label>
                <Input
                  id="template"
                  type="file"
                  accept=".xlsx,.xlsm,.xls"
                  onChange={(e) => setTemplateFile(e.target.files?.[0] || null)}
                />
                <p className="text-xs text-slate-500">{templateFile ? `Selected: ${templateFile.name}` : "No file selected"}</p>
              </div>
              <div className="grid gap-2">
                <Label htmlFor="source">FMECA source workbook</Label>
                <Input
                  id="source"
                  type="file"
                  accept=".xlsx,.xlsm,.xls"
                  onChange={(e) => setSourceFile(e.target.files?.[0] || null)}
                />
                <p className="text-xs text-slate-500">{sourceFile ? `Selected: ${sourceFile.name}` : "No file selected"}</p>
              </div>
            </div>

            <div className="flex flex-wrap gap-3">
              <Button onClick={handleGenerate} disabled={!canGenerate} className="rounded-2xl">
                {isGenerating ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Upload className="mr-2 h-4 w-4" />}
                {isGenerating ? "Generating..." : "Generate workbook"}
              </Button>
            </div>

            <Alert>
              <AlertDescription>{status}</AlertDescription>
            </Alert>

            {error ? (
              <Alert className="border-red-300 bg-red-50">
                <AlertDescription>{error}</AlertDescription>
              </Alert>
            ) : null}

            {readyName ? (
              <div className="rounded-2xl border bg-white p-4 text-sm text-slate-700">
                <div className="flex items-center gap-2 font-medium">
                  <Download className="h-4 w-4" />
                  Generated: {readyName}
                </div>
                <p className="mt-2 text-slate-500">Click below to download your generated workbook.</p>
                <div className="mt-4">
                  <Button asChild className="rounded-2xl">
                    <a href={downloadUrl} download={readyName}>
                      <Download className="mr-2 h-4 w-4" />
                      Download workbook
                    </a>
                  </Button>
                </div>
              </div>
            ) : null}
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
