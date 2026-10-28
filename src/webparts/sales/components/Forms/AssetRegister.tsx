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
import SearchableDropdown from "../Shared/Dropdown";
import DatePickercontrol from "../Shared/DatePickerField";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export interface AssetRegisterProps {
    match: any;
    spHttpClient: SPHttpClient;
    context: any;
    siteURL: string;
    spContext: any;
    webAbsoluteURL: string;
    currPlantTitle: string;
}

const AssetRegister: React.FC<AssetRegisterProps> = (props) => {
    const listName = "Asset Register";
    const sp = spfi().using(SPFx(props.context));
    const navigate = useNavigate();
    const { id } = useParams<{ id: string }>();
    const [itemId, setItemId] = useState(0);
    const editLoaded = useRef(false);

    const assetNumberRef = useRef<HTMLInputElement>(null);
    const assetDescriptionRef = useRef<HTMLTextAreaElement>(null);
    const makeRef = useRef<HTMLInputElement>(null);
    const modelRef = useRef<HTMLInputElement>(null);
    const serialNumberRef = useRef<HTMLInputElement>(null);
    const processorRef = useRef<HTMLInputElement>(null);
    const hddRef = useRef<HTMLInputElement>(null);
    const ssdRef = useRef<HTMLInputElement>(null);
    const monitorMakeRef = useRef<HTMLInputElement>(null);
    const monitorModelRef = useRef<HTMLInputElement>(null);
    const monitorSizeRef = useRef<HTMLInputElement>(null);
    const monitorSerialRef = useRef<HTMLInputElement>(null);
    const describeProblemRef = useRef<HTMLInputElement>(null);
    const remarksRef = useRef<HTMLTextAreaElement>(null);

    const dateofissueRef = useRef<any>(null);
    const takenbackdateRef = useRef<any>(null);
    const reissuedateRef = useRef<any>(null);
    const correcteddateRef = useRef<any>(null);

    const [formdata, setFormData] = useState({
        AssetNo: "",
        AssetDescription: "",
        Asset_x0020_Type: "Hardware",
        Make: "",
        Model: "",
        SerialNumber: "",
        Processor: "",
        RAM: "",
        HDD: "",
        SSD: "",
        MonitorMake: "",
        Monitor_x0020_Model_x0020_Number: "",
        Monitor_x0020_screensize: "",
        Monitor_x0020_SerialNumber: "",
        Date_x0020_of_x0020_issue: "" as any,
        User_x0020_nameId: [] as number[],
        User_x0020_nameEmail: "" as string,
        Taken_x0020_back_x0020_date: "" as any,
        Re_x0020_issue_x0020_date: "" as any,
        Re_x0020_issued_x0020_toId: [] as number[],
        Re_x0020_issued_x0020_toEmail: "" as string,
        Working: "Yes",
        Remarks: "",
        Describe_x0020_problem: "",
        Corrected_x0020_date: "" as any
    });

    // const isHardware = formdata.Asset_x0020_Type === "Hardware";
    const isNotWorking = formdata.Working === "No";

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
                    "User_x0020_name/Id", "User_x0020_name/EMail",
                    "Re_x0020_issued_x0020_to/Id", "Re_x0020_issued_x0020_to/EMail"
                )
                .expand("User_x0020_name", "Re_x0020_issued_x0020_to")();

            setFormData({
                AssetNo: item.AssetNo || "",
                AssetDescription: item.AssetDescription || "",
                Asset_x0020_Type: item.Asset_x0020_Type || "Hardware",
                Make: item.Make || "",
                Model: item.Model || "",
                SerialNumber: item.SerialNumber || "",
                Processor: item.Processor || "",
                RAM: item.RAM || "",
                HDD: item.HDD || "",
                SSD: item.SSD || "",
                MonitorMake: item.MonitorMake || "",
                Monitor_x0020_Model_x0020_Number: item.Monitor_x0020_Model_x0020_Number || "",
                Monitor_x0020_screensize: item.Monitor_x0020_screensize?.toString() || "",
                Monitor_x0020_SerialNumber: item.Monitor_x0020_SerialNumber || "",
                Date_x0020_of_x0020_issue: item.Date_x0020_of_x0020_issue || "",
                User_x0020_nameId: item.User_x0020_nameId ? [item.User_x0020_nameId] : [],
                User_x0020_nameEmail: item.User_x0020_name?.EMail || "",
                Taken_x0020_back_x0020_date: item.Taken_x0020_back_x0020_date || "",
                Re_x0020_issue_x0020_date: item.Re_x0020_issue_x0020_date || "",
                Re_x0020_issued_x0020_toId: item.Re_x0020_issued_x0020_toId ? [item.Re_x0020_issued_x0020_toId] : [],
                Re_x0020_issued_x0020_toEmail: item.Re_x0020_issued_x0020_to?.EMail || "",
                Working: item.Working === false ? "No" : "Yes",
                Remarks: item.Remarks || "",
                Describe_x0020_problem: item.Describe_x0020_problem || "",
                Corrected_x0020_date: item.Corrected_x0020_date || ""
            });
        } catch (e) {
            console.log(e);
            showToast("error", "Error loading edit data");
        } finally {
            hideLoader();
        }
    };

    /* ---------------- CREATE / UPDATE ---------------- */
    const insertOrUpdate = async () => {
        try {
            const data: any = {
                AssetNo: formdata.AssetNo,
                AssetDescription: formdata.AssetDescription,
                Asset_x0020_Type: formdata.Asset_x0020_Type,
                Make: formdata.Make,
                Model: formdata.Model,
                SerialNumber: formdata.SerialNumber,
                Processor:  formdata.Processor,
                RAM:  formdata.RAM ,
                HDD:  formdata.HDD ,
                SSD:  formdata.SSD ,
                MonitorMake: formdata.MonitorMake,
                Monitor_x0020_Model_x0020_Number:formdata.Monitor_x0020_Model_x0020_Number,
                Monitor_x0020_screensize: (parseFloat(formdata.Monitor_x0020_screensize) || 0) ,
                Monitor_x0020_SerialNumber:formdata.Monitor_x0020_SerialNumber,
                Date_x0020_of_x0020_issue: formdata.Date_x0020_of_x0020_issue || null,
                User_x0020_nameId: formdata.User_x0020_nameId.length > 0 ? formdata.User_x0020_nameId[0] : null,
                Taken_x0020_back_x0020_date: formdata.Taken_x0020_back_x0020_date || null,
                Re_x0020_issue_x0020_date: formdata.Re_x0020_issue_x0020_date || null,
                Re_x0020_issued_x0020_toId: formdata.Re_x0020_issued_x0020_toId.length > 0 ? formdata.Re_x0020_issued_x0020_toId[0] : null,
                Working: formdata.Working === "Yes",   // Yes/No column = boolean
                Remarks: formdata.Remarks,
                Describe_x0020_problem: isNotWorking ? formdata.Describe_x0020_problem : "",
                Corrected_x0020_date: isNotWorking ? (formdata.Corrected_x0020_date || null) : null
            };

            if (itemId > 0) {
                await sp.web.lists.getByTitle(listName).items.getById(itemId).update(data);
            } else {
                await sp.web.lists.getByTitle(listName).items.add(data);
            }

            showToast("success", itemId > 0 ? "Asset updated successfully" : "Asset submitted successfully");
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
    //         await sp.web.lists.getByTitle(listName).items.getById(deleteId).delete();
    //         showToast("success", "Asset deleted successfully");
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

    const handleChangeDropdown = (selectedOption: any, actionMeta: any) => {
        if (!actionMeta?.name) return;
        const value = actionMeta.action === "clear" ? "" : selectedOption?.value;
        setFormData(prev => ({ ...prev, [actionMeta.name]: value }));
    };

    const handleDateChange = (dateValue: any, name: string) => {
        setFormData(prev => ({ ...prev, [name]: dateValue ? new Date(dateValue) : null }));
    };

    const handleUserChange = (items: any[]) => {
        const ids = items?.map(item => Number(item.id)) || [];
        setFormData(prev => ({ ...prev, User_x0020_nameId: ids }));
    };

    const handleReissuedToChange = (items: any[]) => {
        const ids = items?.map(item => Number(item.id)) || [];
        setFormData(prev => ({ ...prev, Re_x0020_issued_x0020_toId: ids }));
    };

    const handleSubmit = async (event: any) => {
        event.preventDefault();
        showLoader();

        const data = {
            AssetNo: { val: formdata.AssetNo, required: true, Name: "Asset no", Type: ControlType.string, Focusid: assetNumberRef },
            AssetDescription: { val: formdata.AssetDescription, required: true, Name: "Description", Type: ControlType.string, Focusid: assetDescriptionRef },
            Asset_x0020_Type: { val: formdata.Asset_x0020_Type, required: true, Name: "Asset type", Type: ControlType.reactSelect, Focusid: "txtAssetType" },
            Make: { val: formdata.Make, required: true, Name: "Make", Type: ControlType.string, Focusid: makeRef },
            Model: { val: formdata.Model, required: true, Name: "Model", Type: ControlType.string, Focusid: modelRef },
            SerialNumber: { val: formdata.SerialNumber, required: true, Name: "Serial no", Type: ControlType.string, Focusid: serialNumberRef },

            Processor: { val: formdata.Processor, required:true, Name: "Processor", Type: ControlType.string, Focusid: processorRef },
            RAM: { val: formdata.RAM, required: true, Name: "RAM", Type: ControlType.reactSelect, Focusid: "txtRAM" },
            HDD: { val: formdata.HDD, required: true, Name: "HDD", Type: ControlType.string, Focusid: hddRef },
            SSD: { val: formdata.SSD, required: true, Name: "SSD", Type: ControlType.string, Focusid: ssdRef },
            MonitorMake: { val: formdata.MonitorMake, required: true, Name: "Monitor make", Type: ControlType.string, Focusid: monitorMakeRef },
            MonitorModel: { val: formdata.Monitor_x0020_Model_x0020_Number, required: true, Name: "Monitor model no", Type: ControlType.string, Focusid: monitorModelRef },
            MonitorSize: { val: formdata.Monitor_x0020_screensize, required: true, Name: "Screen size", Type: ControlType.string, Focusid: monitorSizeRef },
            MonitorSerial: { val: formdata.Monitor_x0020_SerialNumber, required: true, Name: "Monitor serial no", Type: ControlType.string, Focusid: monitorSerialRef },

            DateOfIssue: { val: formdata.Date_x0020_of_x0020_issue, required: true, Name: "Date of issue", Type: ControlType.date, Focusid: "dtDate_x0020_of_x0020_issue" },
            UserName: { val: formdata.User_x0020_nameId, required: true, Name: "User name", Type: ControlType.people, Focusid: "userNameDiv" },
            TakenBackDate: { val: formdata.Taken_x0020_back_x0020_date, required: true, Name: "Taken back date", Type: ControlType.date, Focusid: "dtTaken_x0020_back_x0020_date" },
            ReissueDate: { val: formdata.Re_x0020_issue_x0020_date, required: true, Name: "Re-issue date", Type: ControlType.date, Focusid: "dtRe_x0020_issue_x0020_date" },
            ReissuedTo: { val: formdata.Re_x0020_issued_x0020_toId, required: true, Name: "Re-issued to", Type: ControlType.people, Focusid: "reissuedToDiv" },

            Working: { val: formdata.Working, required: true, Name: "Working status", Type: ControlType.reactSelect, Focusid: "txtWorking" },
            DescribeProblem: { val: formdata.Describe_x0020_problem, required: isNotWorking, Name: "Describe problem", Type: ControlType.string, Focusid: describeProblemRef },
            CorrectedDate: { val: formdata.Corrected_x0020_date, required: isNotWorking, Name: "Corrected date", Type: ControlType.date, Focusid: "Corrected_x0020_date" },
            Remarks: { val: formdata.Remarks, required: true, Name: "Remarks", Type: ControlType.string, Focusid: remarksRef }
        };

        const isValid = formValidation.FormValidation(data);
        if (!isValid.status) {
            showToast("error", isValid.message);
            hideLoader();
            return;
        }

        insertOrUpdate();
    };

    const closeForm = () => {
        navigate("/AssetRegisterView"); // change to your view route
    };

    const onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    };

    /* ---------------- UI ---------------- */
    return (
        <div className="container-fluid">
            <div className="light-box border-box-shadow">
                <div className="div-form-title">
                    <div className="form-title">
                        Asset Register{itemId > 0 ? ` - (${formdata.AssetNo})` : ""}
                    </div>
                    <span className="span-mandatory-text">
                        <span className="text-danger">*</span> are mandatory fields
                    </span>
                </div>

                <div className="p-2 mx-1 ViewTable">
                    <div className="form-border-box p-2 mx-1 my-2">
                        <div className="row">
                            <div className="col-md-3">
                                <div className="light-text">
                                    <label>Asset no<span className="text-danger">*</span></label>
                                    <input type="text" value={formdata.AssetNo} onChange={handleChange} name="AssetNo" id="txtAssetNo" ref={assetNumberRef} className="form-control" />
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <label>Description<span className="text-danger">*</span></label>
                                    <textarea value={formdata.AssetDescription} onChange={handleChange} name="AssetDescription" id="txtAssetDescription" ref={assetDescriptionRef} className="form-control" rows={2} />
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <SearchableDropdown
                                        label="Asset type"
                                        name="Asset_x0020_Type"
                                        selectedValue={formdata.Asset_x0020_Type}
                                        OptionsList={[{ label: "Hardware", value: "Hardware" }]}
                                        OnChange={handleChangeDropdown}
                                        isRequired={true} Title={"Asset Type"} id={"txtAssetType"} className={""} disabled={false} />
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <label>Make<span className="text-danger">*</span></label>
                                    <input type="text" value={formdata.Make} onChange={handleChange} name="Make" id="txtMake" ref={makeRef} className="form-control" />
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <label>Model<span className="text-danger">*</span></label>
                                    <input type="text" value={formdata.Model} onChange={handleChange} name="Model" id="txtModel" ref={modelRef} className="form-control" />
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <label>Serial no<span className="text-danger">*</span></label>
                                    <input type="text" value={formdata.SerialNumber} onChange={handleChange} name="SerialNumber" id="txtSerialNumber" ref={serialNumberRef} className="form-control" />
                                </div>
                            </div>

                        
                                    <h4>Configuration</h4>
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label>Processor<span className="text-danger">*</span></label>
                                            <input type="text" value={formdata.Processor} onChange={handleChange} name="Processor" id="txtProcessor" ref={processorRef} className="form-control" />
                                        </div>
                                    </div>
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <SearchableDropdown
                                                label="RAM"
                                                name="RAM"
                                                selectedValue={formdata.RAM}
                                                OptionsList={["4GB", "8GB", "16GB", "32GB", "64GB"].map(v => ({ label: v, value: v }))}
                                                OnChange={handleChangeDropdown}
                                                isRequired={true} Title={"RAM"} id={"txtRAM"} className={""} disabled={false} />
                                        </div>
                                    </div>
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label>HDD<span className="text-danger">*</span></label>
                                            <input type="text" value={formdata.HDD} onChange={handleChange} name="HDD" id="txtHDD" ref={hddRef} className="form-control" />
                                        </div>
                                    </div>
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label>SSD<span className="text-danger">*</span></label>
                                            <input type="text" value={formdata.SSD} onChange={handleChange} name="SSD" id="txtSSD" ref={ssdRef} className="form-control" />
                                        </div>
                                    </div>

                                    <h4>Monitor details</h4>
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label>Make<span className="text-danger">*</span></label>
                                            <input type="text" value={formdata.MonitorMake} onChange={handleChange} name="MonitorMake" id="txtMonitorMake" ref={monitorMakeRef} className="form-control" />
                                        </div>
                                    </div>
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label>Model no<span className="text-danger">*</span></label>
                                            <input type="text" value={formdata.Monitor_x0020_Model_x0020_Number} onChange={handleChange} name="Monitor_x0020_Model_x0020_Number" id="txtMonitorModelNumber" ref={monitorModelRef} className="form-control" />
                                        </div>
                                    </div>
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label>Screen size (inches)<span className="text-danger">*</span></label>
                                            <input type="text" value={formdata.Monitor_x0020_screensize} onChange={handleChange} name="Monitor_x0020_screensize" id="txtMonitorScreenSize" ref={monitorSizeRef} className="form-control" />
                                        </div>
                                    </div>
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label>Serial no<span className="text-danger">*</span></label>
                                            <input type="text" value={formdata.Monitor_x0020_SerialNumber} onChange={handleChange} name="Monitor_x0020_SerialNumber" id="txtMSerialNumber" ref={monitorSerialRef} className="form-control" />
                                        </div>
                                    </div>
                              
                            <h4>Issue details</h4>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <div ref={dateofissueRef}>
                                        <label htmlFor="dtDate_x0020_of_x0020_issue">Date of issue <span className="text-danger">*</span></label>
                                        <div className="custom-datepicker" id="divDateOfIssue">
                                            <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={formdata.Date_x0020_of_x0020_issue} ref={dateofissueRef} title={formdata.Date_x0020_of_x0020_issue} id="dtDate_x0020_of_x0020_issue" startDate={undefined} endDate={new Date()} name="Date_x0020_of_x0020_issue" onDatechange={(p: any) => handleDateChange(p[0], "Date_x0020_of_x0020_issue")} highlightDate={new Date()} showIcon />
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text" id="userNameDiv">
                                    <label>User name <span className="text-danger">*</span></label>
                                    <PeoplePicker
                                        context={props.context}
                                        webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
                                        titleText=""
                                        personSelectionLimit={1}
                                        showtooltip={true}
                                        required={true}
                                        disabled={false}
                                        ensureUser={true}
                                        principalTypes={[PrincipalType.User]}
                                        resolveDelay={500}
                                        onChange={handleUserChange}
                                        defaultSelectedUsers={formdata.User_x0020_nameEmail ? [formdata.User_x0020_nameEmail] : []}
                                    />
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <div ref={takenbackdateRef}>
                                        <label htmlFor="dtTaken_x0020_back_x0020_date">Taken back date <span className="text-danger">*</span></label>
                                        <div className="custom-datepicker" id="divTakenBackDate">
                                            <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={formdata.Taken_x0020_back_x0020_date} ref={takenbackdateRef} title={formdata.Taken_x0020_back_x0020_date} id="dtTaken_x0020_back_x0020_date" startDate={undefined} endDate={new Date()} name="Taken_x0020_back_x0020_date" onDatechange={(p: any) => handleDateChange(p[0], "Taken_x0020_back_x0020_date")} highlightDate={new Date()} showIcon />
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <div ref={reissuedateRef}>
                                        <label htmlFor="dtRe_x0020_issue_x0020_date">Re-issue date <span className="text-danger">*</span></label>
                                        <div className="custom-datepicker" id="divReissueDate">
                                            <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={formdata.Re_x0020_issue_x0020_date} ref={reissuedateRef} title={formdata.Re_x0020_issue_x0020_date} id="dtRe_x0020_issue_x0020_date" startDate={undefined} endDate={new Date()} name="Re_x0020_issue_x0020_date" onDatechange={(p: any) => handleDateChange(p[0], "Re_x0020_issue_x0020_date")} highlightDate={new Date()} showIcon />
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div className="col-md-3">
                                <div className="light-text" id="reissuedToDiv">
                                    <label>Re-issued to <span className="text-danger">*</span></label>
                                    <PeoplePicker
                                        context={props.context}
                                        webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
                                        titleText=""
                                        personSelectionLimit={1}
                                        showtooltip={true}
                                        required={true}
                                        disabled={false}
                                        ensureUser={true}
                                        principalTypes={[PrincipalType.User]}
                                        resolveDelay={500}
                                        onChange={handleReissuedToChange}
                                        defaultSelectedUsers={formdata.Re_x0020_issued_x0020_toEmail ? [formdata.Re_x0020_issued_x0020_toEmail] : []}
                                    />
                                </div>
                            </div>
                        
                            <h4>Working status</h4>
                            <div className="col-md-3">
                                <div className="light-text">
                                    <SearchableDropdown
                                        label="Working status"
                                        name="Working"
                                        selectedValue={formdata.Working}
                                        OptionsList={[{ label: "Yes", value: "Yes" }, { label: "No", value: "No" }]}
                                        OnChange={handleChangeDropdown}
                                        isRequired={true} Title={"Working"} id={"txtWorking"} className={""} disabled={false} />
                                </div>
                            </div>
                            {isNotWorking && (
                                <>
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <label>Describe problem<span className="text-danger">*</span></label>
                                            <input type="text" value={formdata.Describe_x0020_problem} onChange={handleChange} name="Describe_x0020_problem" id="txtDescribeProblem" ref={describeProblemRef} className="form-control" />
                                        </div>
                                    </div>
                                    <div className="col-md-3">
                                        <div className="light-text">
                                            <div ref={correcteddateRef}>
                                                <label htmlFor="Corrected_x0020_date">Corrected date<span className="text-danger">*</span></label>
                                                <div className="custom-datepicker" id="divCorrectedDate">
                                                    <DatePickercontrol placeholder="MM/DD/YYYY" selectedDate={formdata.Corrected_x0020_date} ref={correcteddateRef} title={formdata.Corrected_x0020_date} id="Corrected_x0020_date" startDate={undefined} endDate={new Date()} name="Corrected_x0020_date" onDatechange={(p: any) => handleDateChange(p[0], "Corrected_x0020_date")} highlightDate={new Date()} showIcon />
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </>
                            )}

                            <div className="col-md-12">
                                <div className="light-text">
                                    <label>Remarks<span className="text-danger">*</span></label>
                                    <textarea
                                        name="Remarks"
                                        value={formdata.Remarks || ""}
                                        onChange={handleChange}
                                        ref={remarksRef}
                                        className="form-control"
                                        rows={4}
                                        placeholder="Enter Remarks..."
                                    />
                                </div>
                            </div>

                            <div className="col-md-12 py-2 text-center">
                                <button className="btn btn-primary mx-2" onClick={handleSubmit}>
                                    {itemId > 0 ? "Update" : "Submit"}
                                </button>
                                <button className="btn btn-secondary" onClick={closeForm}>
                                    Cancel
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default AssetRegister;