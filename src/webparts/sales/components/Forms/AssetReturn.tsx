import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { useNavigate, useParams } from "react-router-dom";
import { hideLoader, showLoader } from "../Shared/Loader";
import { showToast } from "../Shared/Toaster";
import formValidation from "../Utilities/FormValidator";
import { ActionStatus, ControlType } from "../Constants/Contants";
import DatePickercontrol from "../Shared/DatePickerField";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export interface AssetIssueReturnProps {
    match: any;
    spHttpClient: SPHttpClient;
    context: any;
    siteURL: string;
    spContext: any;
    webAbsoluteURL: string;
    currPlantTitle: string;
}

interface ItemRow {
    Id: number;
    SNo: number;
    Description: string;
    AssetNo: string;
    Remarks: string;
    IssueDate: any;
    IssuedById: number[];
    IssuedByEmail: string;
    ReturnedDate: any;
    ReturnedToId: number[];
    ReturnedToEmail: string;
}

const defaultDescriptions = [
    "LAPTOP", "Charger", "Mouse", "Monitor", "Key Board", "Headset", "Dongle",
    "ID Card", "Access Card", "E-mail Access", "CPU - Desktop", "Mobile (handset)",
    "Mobile SIM card", ""
];

const buildDefaultRows = (): ItemRow[] =>
    defaultDescriptions.map((d, i) => ({
        Id: 0,
        SNo: i + 1,
        Description: d,
        AssetNo: "",
        Remarks: "",
        IssueDate: "",
        IssuedById: [],
        IssuedByEmail: "",
        ReturnedDate: "",
        ReturnedToId: [],
        ReturnedToEmail: ""
    }));

const AssetIssueReturn: React.FC<AssetIssueReturnProps> = (props) => {
    const listName = "Asset Issue Return";
    const itemsListName = "Asset Issue Return Items";
    const sp = spfi().using(SPFx(props.context));
    const navigate = useNavigate();
    const { id } = useParams<{ id: string }>();
    const [itemId, setItemId] = useState(0);
    const [loadKey, setLoadKey] = useState(0); // re-mounts people pickers after edit data loads
    const editLoaded = useRef(false);

    const employeeNameRef = useRef<HTMLInputElement>(null);
    const employeeNoRef = useRef<HTMLInputElement>(null);
    const locationRef = useRef<HTMLInputElement>(null);
    const deptNameRef = useRef<HTMLInputElement>(null);
    const designationRef = useRef<HTMLInputElement>(null);
    const dateRef = useRef<any>(null);

    const [formdata, setFormData] = useState({
        EmployeeName: "",
        EmployeeNo: "",
        Location: "",
        DeptName: "",
        Designation: "",
        DateOfJoining: "" as any,
        DateOfLeaving: "" as any,
        HRRepresentativeId: [] as number[],
        HRRepresentativeEmail: "",
        HRSignDate: "" as any,
        EmployeeSignDate: "" as any,
        ITAdminId: [] as number[],
        ITAdminEmail: "",
        ITSignDate: "" as any,
        ReturnAcknowledgedById: [] as number[],
        ReturnAcknowledgedByEmail: "",
        ReturnAckDate: "" as any
    });

    const [rows, setRows] = useState<ItemRow[]>(buildDefaultRows());

    /* ---------------- READ (edit mode) ---------------- */
    useEffect(() => {
        if (id && !editLoaded.current) {
            editLoaded.current = true;
            const editId = parseInt(id);
            setItemId(editId);
            loadEditData(editId);
        }
    }, [id]);

    const loadEditData = async (editId: number) => {
        try {
            showLoader();
            const item: any = await sp.web.lists
                .getByTitle(listName)
                .items
                .getById(editId)
                .select(
                    "*",
                    "HRRepresentative/Id", "HRRepresentative/EMail",
                    "ITAdmin/Id", "ITAdmin/EMail",
                    "ReturnAcknowledgedBy/Id", "ReturnAcknowledgedBy/EMail"
                )
                .expand("HRRepresentative", "ITAdmin", "ReturnAcknowledgedBy")();

            const children: any[] = await sp.web.lists
                .getByTitle(itemsListName)
                .items
                .select(
                    "*",
                    "IssuedBy/Id", "IssuedBy/EMail",
                    "ReturnedTo/Id", "ReturnedTo/EMail"
                )
                .expand("IssuedBy", "ReturnedTo")
                .filter(`RecordID eq ${editId}`)
                .orderBy("SNo", true)();

            setFormData({
                EmployeeName: item.EmployeeName || "",
                EmployeeNo: item.EmployeeNo || "",
                Location: item.Location || "",
                DeptName: item.DeptName || "",
                Designation: item.Designation || "",
                DateOfJoining: item.DateOfJoining || "",
                DateOfLeaving: item.DateOfLeaving || "",
                HRRepresentativeId: item.HRRepresentativeId ? [item.HRRepresentativeId] : [],
                HRRepresentativeEmail: item.HRRepresentative?.EMail || "",
                HRSignDate: item.HRSignDate || "",
                EmployeeSignDate: item.EmployeeSignDate || "",
                ITAdminId: item.ITAdminId ? [item.ITAdminId] : [],
                ITAdminEmail: item.ITAdmin?.EMail || "",
                ITSignDate: item.ITSignDate || "",
                ReturnAcknowledgedById: item.ReturnAcknowledgedById ? [item.ReturnAcknowledgedById] : [],
                ReturnAcknowledgedByEmail: item.ReturnAcknowledgedBy?.EMail || "",
                ReturnAckDate: item.ReturnAckDate || ""
            });

            // merge saved rows over the default 14 rows (matched by SNo)
            const merged = buildDefaultRows();
            children.forEach((c: any) => {
                const idx = merged.findIndex(r => r.SNo === c.SNo);
                const row: ItemRow = {
                    Id: c.Id,
                    SNo: c.SNo,
                    Description: c.Description || "",
                    AssetNo: c.AssetNo || "",
                    Remarks: c.Remarks || "",
                    IssueDate: c.IssueDate || "",
                    IssuedById: c.IssuedById ? [c.IssuedById] : [],
                    IssuedByEmail: c.IssuedBy?.EMail || "",
                    ReturnedDate: c.ReturnedDate || "",
                    ReturnedToId: c.ReturnedToId ? [c.ReturnedToId] : [],
                    ReturnedToEmail: c.ReturnedTo?.EMail || ""
                };
                if (idx >= 0) merged[idx] = row;
                else merged.push(row);
            });
            setRows(merged);
            setLoadKey(k => k + 1);
        } catch (e) {
            console.log(e);
            showToast("error", "Error loading edit data");
        } finally {
            hideLoader();
        }
    };

    /* ---------------- CREATE / UPDATE ---------------- */
    const isRowUsed = (r: ItemRow) => r.AssetNo.trim() !== "";

    const saveItems = async (parentId: number) => {
        const list = sp.web.lists.getByTitle(itemsListName);

        for (const r of rows) {
            const payload: any = {
                RecordID: parentId,
                SNo: r.SNo,
                Description: r.Description,
                AssetNo: r.AssetNo,
                Remarks: r.Remarks,
                IssueDate: r.IssueDate || null,
                IssuedById: r.IssuedById.length > 0 ? r.IssuedById[0] : null,
                ReturnedDate: r.ReturnedDate || null,
                ReturnedToId: r.ReturnedToId.length > 0 ? r.ReturnedToId[0] : null
            };

            if (isRowUsed(r)) {
                if (r.Id > 0) await list.items.getById(r.Id).update(payload);
                else await list.items.add(payload);
            } else if (r.Id > 0) {
                // row was cleared in edit mode -> delete it
                await list.items.getById(r.Id).delete();
            }
        }
    };

    const insertOrUpdate = async () => {
        try {
            const data: any = {
                EmployeeName: formdata.EmployeeName,
                EmployeeNo: formdata.EmployeeNo,
                Location: formdata.Location,
                DeptName: formdata.DeptName,
                Designation: formdata.Designation,
                DateOfJoining: formdata.DateOfJoining || null,
                DateOfLeaving: formdata.DateOfLeaving || null,
                HRRepresentativeId: formdata.HRRepresentativeId.length > 0 ? formdata.HRRepresentativeId[0] : null,
                HRSignDate: formdata.HRSignDate || null,
                EmployeeSignDate: formdata.EmployeeSignDate || null,
                ITAdminId: formdata.ITAdminId.length > 0 ? formdata.ITAdminId[0] : null,
                ITSignDate: formdata.ITSignDate || null,
                ReturnAcknowledgedById: formdata.ReturnAcknowledgedById.length > 0 ? formdata.ReturnAcknowledgedById[0] : null,
                ReturnAckDate: formdata.ReturnAckDate || null
            };

            let parentId = itemId;

            if (itemId > 0) {
                await sp.web.lists.getByTitle(listName).items.getById(itemId).update(data);
            } else {
                const result: any = await sp.web.lists.getByTitle(listName).items.add(data);
                parentId = result.Id ?? result.data?.Id;
            }

            await saveItems(parentId);

            showToast("success", itemId > 0 ? "Acknowledgement updated successfully" : "Acknowledgement submitted successfully");
            closeForm();
        } catch (e) {
            console.log(e);
            onError();
        } finally {
            hideLoader();
        }
    };

    /* ---------------- DELETE (call from the view screen) ---------------- */
    // const deleteItem = async (deleteId: number) => {
    //     try {
    //         showLoader();
    //         const children: any[] = await sp.web.lists
    //             .getByTitle(itemsListName).items
    //             .select("Id").filter(`RecordID eq ${deleteId}`)();
    //         for (const c of children) {
    //             await sp.web.lists.getByTitle(itemsListName).items.getById(c.Id).delete();
    //         }
    //         await sp.web.lists.getByTitle(listName).items.getById(deleteId).delete();
    //         showToast("success", "Acknowledgement deleted successfully");
    //     } catch (e) {
    //         console.log(e);
    //         onError();
    //     } finally {
    //         hideLoader();
    //     }
    // };

    /* ---------------- HANDLERS ---------------- */
    const handleChange = (event: any) => {
        const { name, value } = event.target;
        setFormData(prev => ({ ...prev, [name]: value }));
    };

    const handleDateChange = (dateValue: any, name: string) => {
        setFormData(prev => ({ ...prev, [name]: dateValue ? new Date(dateValue) : null }));
    };

    const handlePeopleChange = (items: any[], name: string) => {
        const ids = items?.map(item => Number(item.id)) || [];
        setFormData(prev => ({ ...prev, [name]: ids }));
    };

    const updateRow = (index: number, field: keyof ItemRow, value: any) => {
        setRows(prev => prev.map((r, i) => (i === index ? { ...r, [field]: value } : r)));
    };

    const handleRowDateChange = (index: number, field: "IssueDate" | "ReturnedDate", dateValue: any) => {
        updateRow(index, field, dateValue ? new Date(dateValue) : null);
    };

    const handleRowPeopleChange = (index: number, field: "IssuedById" | "ReturnedToId", items: any[]) => {
        updateRow(index, field, items?.map(item => Number(item.id)) || []);
    };

    const toDate = (v: any) => (v ? new Date(v).getTime() : 0);

    const validateRows = (): string => {
        const used = rows.filter(isRowUsed);
        if (used.length === 0) return "Please enter at least one asset (Asset No).";

        for (const r of used) {
            if (!r.Description.trim()) return `Row ${r.SNo}: Description is required.`;
            if (!r.IssueDate) return `Row ${r.SNo}: Issue Date is required.`;
            if (r.IssuedById.length === 0) return `Row ${r.SNo}: Issued By is required.`;
            if (r.ReturnedDate && r.ReturnedToId.length === 0) return `Row ${r.SNo}: Returned To is required when Returned Date is entered.`;
            if (r.ReturnedToId.length > 0 && !r.ReturnedDate) return `Row ${r.SNo}: Returned Date is required when Returned To is selected.`;
            if (r.ReturnedDate && toDate(r.ReturnedDate) < toDate(r.IssueDate)) return `Row ${r.SNo}: Returned Date cannot be before Issue Date.`;
        }
        return "";
    };

    const handleSubmit = async (event: any) => {
        event.preventDefault();
        showLoader();

        const data = {
            EmployeeName: { val: formdata.EmployeeName, required: true, Name: "Employee Name", Type: ControlType.string, Focusid: employeeNameRef },
            EmployeeNo: { val: formdata.EmployeeNo, required: true, Name: "Employee No", Type: ControlType.string, Focusid: employeeNoRef },
            Location: { val: formdata.Location, required: true, Name: "Location", Type: ControlType.string, Focusid: locationRef },
            DeptName: { val: formdata.DeptName, required: true, Name: "Dept Name", Type: ControlType.string, Focusid: deptNameRef },
            Designation: { val: formdata.Designation, required: true, Name: "Designation", Type: ControlType.string, Focusid: designationRef },
            DateOfJoining: { val: formdata.DateOfJoining, required: true, Name: "Date of Joining", Type: ControlType.date, Focusid: "dtDateOfJoining" }
        };

        const isValid = formValidation.FormValidation(data);
        if (!isValid.status) {
            showToast("error", isValid.message);
            hideLoader();
            return;
        }

        if (formdata.DateOfLeaving && toDate(formdata.DateOfLeaving) < toDate(formdata.DateOfJoining)) {
            showToast("error", "Date of Leaving cannot be before Date of Joining.");
            hideLoader();
            return;
        }

        const rowError = validateRows();
        if (rowError) {
            showToast("error", rowError);
            hideLoader();
            return;
        }

        insertOrUpdate();
    };

    const closeForm = () => {
        navigate("/AssetIssueReturnView"); // change to your view route
    };

    const onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    };

    /* ---------------- UI helpers ---------------- */
    const renderDate = (label: string, fieldId: string, name: string, value: any, required = false, endDate: any = undefined) => (
        <div className="col-md-3">
            <div className="light-text">
                <label htmlFor={fieldId}>{label} {required && <span className="text-danger">*</span>}</label>
                <div className="custom-datepicker" id={`div${fieldId}`}>
                    <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={value} ref={dateRef} title={value} id={fieldId} startDate={undefined} endDate={endDate} name={name} onDatechange={(p: any) => handleDateChange(p[0], name)} highlightDate={new Date()} showIcon />
                </div>
            </div>
        </div>
    );

    const renderPeople = (label: string, divId: string, name: string, email: string) => (
        <div className="col-md-3">
            <div className="light-text" id={divId}>
                <label>{label}</label>
                <PeoplePicker
                    key={`${divId}-${loadKey}`}
                    context={props.context}
                    webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
                    titleText=""
                    personSelectionLimit={1}
                    showtooltip={true}
                    disabled={false}
                    ensureUser={true}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={500}
                    onChange={(items: any[]) => handlePeopleChange(items, name)}
                    defaultSelectedUsers={email ? [email] : []}
                />
            </div>
        </div>
    );

    /* ---------------- UI ---------------- */
    return (
        <div className="container-fluid">
            <div className="light-box border-box-shadow">
                <div className="div-form-title">
                    <div className="form-title">
                        Asset Issue and Return - Acknowledgement Form{itemId > 0 ? ` - (${formdata.EmployeeNo})` : ""}
                    </div>
                    <span className="span-mandatory-text">
                        <span className="text-danger">*</span> are mandatory fields
                    </span>
                </div>

                <div className="p-2 mx-1 ViewTable">
                    <div className="form-border-box p-2 mx-1 my-2">
                        <div className="row">
                            <div className="col-md-12">
                                <div className="light-text">
                                    <label>Employee Name<span className="text-danger">*</span></label>
                                    <input type="text" value={formdata.EmployeeName} onChange={handleChange} name="EmployeeName" id="txtEmployeeName" ref={employeeNameRef} className="form-control" />
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <label>Employee No<span className="text-danger">*</span></label>
                                    <input type="text" value={formdata.EmployeeNo} onChange={handleChange} name="EmployeeNo" id="txtEmployeeNo" ref={employeeNoRef} className="form-control" />
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <label>Location<span className="text-danger">*</span></label>
                                    <input type="text" value={formdata.Location} onChange={handleChange} name="Location" id="txtLocation" ref={locationRef} className="form-control" />
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <label>Dept Name<span className="text-danger">*</span></label>
                                    <input type="text" value={formdata.DeptName} onChange={handleChange} name="DeptName" id="txtDeptName" ref={deptNameRef} className="form-control" />
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <label>Designation<span className="text-danger">*</span></label>
                                    <input type="text" value={formdata.Designation} onChange={handleChange} name="Designation" id="txtDesignation" ref={designationRef} className="form-control" />
                                </div>
                            </div>
                            {renderDate("Date of Joining", "dtDateOfJoining", "DateOfJoining", formdata.DateOfJoining, true, new Date())}
                            {renderDate("Date of Leaving / Last Working Day", "dtDateOfLeaving", "DateOfLeaving", formdata.DateOfLeaving)}

                            <h4 className="mt-3">Assets</h4>
                            <div className="px-3 table-responsive">
                                <table className="table table-bordered quoteTable">
                                    <thead>
                                        <tr>
                                            <th rowSpan={2}>S. No</th>
                                            <th rowSpan={2}>Description</th>
                                            <th rowSpan={2}>Asset No</th>
                                            <th rowSpan={2}>Remarks</th>
                                            <th colSpan={2} className="text-center">Asset Issue</th>
                                            <th colSpan={2} className="text-center">Asset Return</th>
                                        </tr>
                                        <tr>
                                            <th>Issue Date</th>
                                            <th>Issued By</th>
                                            <th>Returned Date</th>
                                            <th>Returned To</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {rows.map((r, i) => (
                                            <tr key={i}>
                                                <td>{r.SNo}</td>
                                                <td>
                                                    <input type="text" className="form-control" value={r.Description} onChange={(e) => updateRow(i, "Description", e.target.value)} />
                                                </td>
                                                <td>
                                                    <input type="text" className="form-control" value={r.AssetNo} onChange={(e) => updateRow(i, "AssetNo", e.target.value)} />
                                                </td>
                                                <td>
                                                    <input type="text" className="form-control" value={r.Remarks} onChange={(e) => updateRow(i, "Remarks", e.target.value)} />
                                                </td>
                                                <td>
                                                    <div className="custom-datepicker">
                                                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={r.IssueDate} ref={dateRef} title={r.IssueDate} id={`dtIssueDate_${i}`} startDate={undefined} endDate={new Date()} name={`IssueDate_${i}`} onDatechange={(p: any) => handleRowDateChange(i, "IssueDate", p[0])} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </td>
                                                <td>
                                                    <PeoplePicker
                                                        key={`issuedBy-${i}-${loadKey}`}
                                                        context={props.context}
                                                        webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
                                                        titleText=""
                                                        personSelectionLimit={1}
                                                        showtooltip={true}
                                                        disabled={false}
                                                        ensureUser={true}
                                                        principalTypes={[PrincipalType.User]}
                                                        resolveDelay={500}
                                                        onChange={(items: any[]) => handleRowPeopleChange(i, "IssuedById", items)}
                                                        defaultSelectedUsers={r.IssuedByEmail ? [r.IssuedByEmail] : []}
                                                    />
                                                </td>
                                                <td>
                                                    <div className="custom-datepicker">
                                                        <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={r.ReturnedDate} ref={dateRef} title={r.ReturnedDate} id={`dtReturnedDate_${i}`} startDate={undefined} endDate={new Date()} name={`ReturnedDate_${i}`} onDatechange={(p: any) => handleRowDateChange(i, "ReturnedDate", p[0])} highlightDate={new Date()} showIcon />
                                                    </div>
                                                </td>
                                                <td>
                                                    <PeoplePicker
                                                        key={`returnedTo-${i}-${loadKey}`}
                                                        context={props.context}
                                                        webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
                                                        titleText=""
                                                        personSelectionLimit={1}
                                                        showtooltip={true}
                                                        disabled={false}
                                                        ensureUser={true}
                                                        principalTypes={[PrincipalType.User]}
                                                        resolveDelay={500}
                                                        onChange={(items: any[]) => handleRowPeopleChange(i, "ReturnedToId", items)}
                                                        defaultSelectedUsers={r.ReturnedToEmail ? [r.ReturnedToEmail] : []}
                                                    />
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>

                            <h4 className="mt-3">Acknowledgement</h4>
                            {renderPeople("HR Representative", "hrRepDiv", "HRRepresentativeId", formdata.HRRepresentativeEmail)}
                            {renderDate("HR Sign Date", "dtHRSignDate", "HRSignDate", formdata.HRSignDate, false, new Date())}
                            {renderDate("Employee Sign Date (Received)", "dtEmployeeSignDate", "EmployeeSignDate", formdata.EmployeeSignDate, false, new Date())}
                            <div className="col-md-3"></div>
                            {renderPeople("IT Admin / Dept Member", "itAdminDiv", "ITAdminId", formdata.ITAdminEmail)}
                            {renderDate("IT Admin Sign Date", "dtITSignDate", "ITSignDate", formdata.ITSignDate, false, new Date())}
                            {renderPeople("Asset Return Acknowledged By", "returnAckDiv", "ReturnAcknowledgedById", formdata.ReturnAcknowledgedByEmail)}
                            {renderDate("Return Acknowledged Date", "dtReturnAckDate", "ReturnAckDate", formdata.ReturnAckDate, false, new Date())}

                            <div className="col-md-12 py-2 text-center">
                                <button className="btn btn-primary mx-2" onClick={handleSubmit}>
                                    {itemId > 0 ? "Update" : "Submit"}
                                </button>
                                <button className="btn btn-secondary" onClick={closeForm}>
                                    Cancel
                                </button>
                            </div>

                            <div className="col-md-12 text-center">
                                <small>Template No: SCSIL-IT-FT-02, Version No: 1.0, Date: 01-Apr-2026</small>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default AssetIssueReturn;