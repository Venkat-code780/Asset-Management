import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
// import TableGenerator from "../Shared/TableGenerator";
import { hideLoader, showLoader } from "../Shared/Loader";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faEdit, faPenToSquare, faPlus } from "@fortawesome/free-solid-svg-icons";
import { highlightCurrentNav } from "../Utilities/HighlightCurrentComponent";
import { ActionStatus, ControlType } from "../Constants/Contants";
import formValidation from "../Utilities/FormValidator";
import { showToast } from "../Shared/Toaster";
import { NavLink } from "react-router-dom";
import SearchableDropdown from "../Shared/Dropdown";
import { initCommonFunctions } from "../Utilities/CommonFunctions";
import InputCheckBox from "../Shared/InputCheckBox";
import AGGridDataTable from "../Shared/AGGridDataTable";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
export interface LocationProps {
    match: any;
    spHttpClient: SPHttpClient;
    context: any;
    siteURL: string;
    spContext: any;
    webAbsoluteURL: string;
    currPlantTitle: string;

}

const Location: React.FC<LocationProps> = (props) => {
    const rootSiteURL = props.spContext.siteAbsoluteUrl;
    const currentSiteURL = props.spContext.webAbsoluteUrl;
    const JvisURL = `${rootSiteURL}/mayco`;
    const { getListItems } = initCommonFunctions(props.context, props.siteURL);


    const listName = "Location";


    const sp = spfi().using(SPFx(props.context));
    // const txtPPEType = useRef<HTMLInputElement>(null);
    const LocationCodeRef = useRef<HTMLInputElement>(null);

    const [actionsData, setActionsData] = useState<any[]>([]);
    // const [pageNumber, setPageNumber] = useState(1);
    const [isFormOpen, setIsFormOpen] = useState(false);
    const [itemId, setItemId] = useState(0);
    const [redirect, setRedirect] = useState(false);
    // const [displayMessage, setDisplayMessage] = useState("");
    const [Location, setLocation] = useState<any[]>([]);
    const [formData, setFormData] = useState({
        Title: "",
    });

    // Load initial data
    useEffect(() => {
        highlightCurrentNav("liPPETypes");
        document.title = "JVIS - Quality | LPA Auditor Levels";
        loadListData();
    }, []);

    useEffect(() => {
        if (redirect) loadListData();
    }, [redirect]);

    // Load Auditors filtered by Plant 'Shelby' and Audit Level

    // Load main table data
    const loadListData = async () => {
        try {
            showLoader();
            setRedirect(false);
            const items = await sp.web.lists
                .getByTitle(listName)
                .items.select('*')
                .top(2000)
                .orderBy("Modified", false)();
            const tableData = items.map((item: any) => ({

                Title: item.Title,
            }));
            setLocation(tableData);
        } catch (e) {
            console.log(e);
            onError();
        } finally {
            hideLoader();
        }
    };

    // Edit item
    const editItem = async (Id: number) => {
        try {
            showLoader();
            setIsFormOpen(true);
            setItemId(Id);
            const item = await sp.web.lists.getByTitle(listName).items.getById(Id).select('*')();
               setFormData({

            Title: item.Title || "",
        });

            setTimeout(() => {
                LocationCodeRef.current?.focus();
            }, 0);
        } catch (e) {
            console.log(e);
            onError();
        } finally {
            hideLoader();
        }
    };

    // Add new item
    const addNew = () => {
        setIsFormOpen(true);
        setItemId(0);
        setFormData({ Title: "" });
        setTimeout(() => {
            LocationCodeRef.current?.focus();

        }, 0);
    };

    // Check duplicate
    const checkDuplicate = async () => {
        try {
            showLoader();
            const filterQuery = `Title eq '${formData.Title}'`;
            const results = await sp.web.lists.getByTitle(listName).items.filter(itemId > 0 ? `${filterQuery} and Id ne ${itemId}` : filterQuery)();
            hideLoader();
            if (results.length > 0) {
                showToast("error", "Record already exists");
                return false;
            }
            return true;
        } catch (e) {
            console.log(e);
            hideLoader();
            onError();
            return false;
        }
    };

    // Insert or update
    const insertOrUpdate = async () => {
        try {
            const data: any = {
                Title: formData.Title.trim(),
            };
            if (itemId > 0) {
                await sp.web.lists.getByTitle(listName).items.getById(itemId).update(data);
                showToast("success", "Location updated successfully");
            } else {
                await sp.web.lists.getByTitle(listName).items.add(data);
                showToast("success", "Location submitted successfully");
            }
            setRedirect(true);
            closeForm();
            loadListData();
        } catch (e) {
            console.log(e);
            onError();
        }
        finally {
            hideLoader();
        }
    };

    const handleChange = (event: any) => {
        const { name, value } = event.target;
        setFormData(prev => ({ ...prev, [name]: value }));
    }

     const handlePeoplePickerChange = (items: any[]) => {

    const userIds =
        items?.map(item => Number(item.id)) || [];

    const userEmails =
        items?.map(item => item.secondaryText) || [];

    setFormData(prev => ({
        ...prev,

        AuthorizedPersonId: userIds,

        AuthorizedPersonOptions: userEmails
    }));
};
    // Form submit
    const handleSubmit = async (event: any) => {
        event.preventDefault();
        showLoader();
        const data = {
            Title: { val: formData.Title.trim(), required: true, Name: "Title", Type: ControlType.string, Focusid: LocationCodeRef },
    

        };
        const isValid = formValidation.FormValidation(data);
        if (isValid.status) {
            const validDuplicate = await checkDuplicate();
            if (validDuplicate) insertOrUpdate();
        } else {
            showToast("error", isValid.message);
            hideLoader();
        }
    };

    const closeForm = () => {
        setIsFormOpen(false);
        setFormData({ Title: ""});
    };
    const onError = () => {
        showToast("error", ActionStatus.Error);
        hideLoader();
    };

   
    const columns = React.useMemo(() => [
        {
            headerName: "Edit",
            field: "Id",
            width: 80,
            minWidth: 80,
            maxWidth: 80,
            sortable: false,
            filter: false,
            cellRenderer: (params: any) => {
                const record = params.data;

                return (
                    <NavLink
                        title="Edit"
                        className="csrLink ms-draggable"
                        to=""
                        onClick={(e) => {
                            e.preventDefault();
                            editItem(record.Id);
                        }}
                    >
                        <FontAwesomeIcon icon={faEdit} />
                    </NavLink>
                );
            },
        },
        {
            headerName: "Title",
            field: "Title",
            sortable: true,
            filter: "agTextColumnFilter",
            resizable: true,
            flex: 1,
            getQuickFilterText: (params: any) => params.value || "",
        },
       
    ], []);

    return (
        <div className="container-fluid">
            <div className="light-box border-box-shadow">
                <div className="div-form-title">
                    <div className="form-title">Location</div>
                    {isFormOpen && (
                        <span className="span-mandatory-text">
                            <span className="text-danger">*</span> are mandatory fields
                        </span>
                    )}
                </div>

                <div className="p-2 mx-1 ViewTable">
                    {!isFormOpen && (
                        <div className="text-end my-2 mx-1">
                            <button className="NewButton btn" onClick={addNew}>
                                <FontAwesomeIcon icon={faPlus} /> New
                            </button>
                        </div>
                    )}

                    {isFormOpen && (
                        <div className="form-border-box p-2 mx-1 my-2">
                            <div className="row">
                                <div className="col-md-3">
                                    <div className="light-text">
                                        <input
                                            className="form-control"
                                            type="text"
                                            name="Title"
                                            value={formData.Title}
                                            onChange={handleChange}
                                            ref={LocationCodeRef}
                                            maxLength={250}
                                        />

                                        <label>
                                            Location <span className="mandatoryhastrick">*</span>
                                        </label>
                                    </div>
                                </div>
                                <div className="col-md-3 py-2 text-center">
                                    <button className="btn btn-primary mx-2" onClick={handleSubmit}>
                                        {itemId ? "Update" : "Submit"}
                                    </button>
                                    <button className="btn btn-secondary" onClick={closeForm}>
                                        Cancel
                                    </button>
                                </div>
                            </div>
                        </div>
                    )}
                    <div className="table-head-1st-td right-search-table mb-3">
                        <AGGridDataTable
                            data={Location}
                            columns={columns}
                            showExportExcel={false}
                            showAddButton={false}
                            customBtnClass="px-1 text-right"
                            btnDivID=""
                            btnSpanID=""
                            btnTitle=""
                            searchBoxLeft={true}
                            onRowClicked={(event: any) => editItem(event.data.Id)}
                            domLayout="normal"
                            suppressColumnVirtualization={true}
                            ensureDomOrder={true}
                            suppressHorizontalScroll={false}
                            suppressSizeToFit={true}
                            suppressColumnHiding={true}
                            suppressAutoSize={true}
                            suppressColumnMoveAnimation={true}
                            suppressMovableColumns={true}
                        />
                    </div>
                </div>
            </div>
        </div>
    );
};

export default Location;