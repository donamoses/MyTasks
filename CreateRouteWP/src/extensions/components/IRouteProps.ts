export interface IRouteProps {
    onClose: () => void;

    isOpen: boolean;
    paneltype: any;
    dismissPanel: () => void;
    itemidprops: number;
    listId: string;
    context: any | null;
    PlannedDateprops: any;
    Districtprops: any;
    DealerNameprops: any;
    ContactNumberprops: any;
    Locationprops: any;
    AssignToprops: any;
    PlannedVisitTimeprops: any;
    Remarksprops: any;
    dealeroptionsprops: any[];
    assigntooptionprops: any[];

}