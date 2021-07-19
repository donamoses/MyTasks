import { WebPartContext } from '@microsoft/sp-webpart-base';
//import { ExtensionContext} from '@microsoft/sp-webpart-base';
import { ExtensionContext } from "@microsoft/sp-extension-base";
export interface ICreateEmployeeProps {
    Nationality: number;
    AlternativeEmail: string;
    AlternativePhone: any;
    WorkPhone: any;
    Country: any;
    City: any;
    PIN: any;
    PassportDetails: any;
    IDProofDetails: any;
    DrivingLicense: any;
    Address: any;
    NoticePeriod: any;
    description: string;
    isOpen: boolean;
    onClose: () => void;
    paneltype: any;
    issetpanel: boolean;
    context: any | null;
    currentTitle: string;
    itemId: number;
    listId: string;
    dismissPanel: () => void;
    empid: any;
    setIsOpen: boolean;
    bloodgp: any;
    id: any;
    birthdate: any;
    joindate: any;
    confirmdate: any;
    mardate: any;
    opttion: any[];
    desgopt: any[];
    locopt: any[];
    state: any[];
    district: any[];
    depthead: any[];
    selectedgender: any;
    selectedmarital: any;
    selecteddept: any;
    selecteddepthead: any;
    selectedlocation: any;
    selecteddesg: any;
    selecteddistrict: any;
    selectedstate: any;
    ReportingOfficer: any;
    permanentaddress: any;
    communicationaddress: any;

    email: any;

    itemid: any;
    Firstname: any;
    Lastname: any;
    PermanentEmployee: boolean;
    Gender: any;
    DateOfbirth: any;
    DateofJoining: any;



    Department: any;
    Designation: any;

    Location: any;

    PermanentAddress: any;
    TemporaryAddress: any;
    MobileNo: any;

    EmailId: any;

    State: any;

    Qualification: any;

    Usernameid: any;



}
