import * as React from 'react';
import { IRouteProps } from './IRouteProps';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownProps, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
import { TextField, DatePicker, DayOfWeek, IDatePickerStrings, mergeStyleSets, DefaultButton, Label, PrimaryButton, DialogFooter, Panel, Spinner, SpinnerType, PanelType, IPanelProps } from "office-ui-fabric-react";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { sp } from "@pnp/sp";
import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import * as moment from 'moment';
export interface IRouteState {
    firstDayOfWeek?: DayOfWeek;
    planneddate: any;
    dealername: any;
    contactnumber: any;
    contactnumbererrormsg: any;
    remarks: any;
    plannedvisittime: any;
    location: any;
    district: any;
    assignto: any;
    dealeroption: any[];
    locationoption: any[];
    assigntooption: any[];
    districtoption: any[];
}
const DayPickerStrings: IDatePickerStrings = {
    months: [
        'January',
        'February',
        'March',
        'April',
        'May',
        'June',
        'July',
        'August',
        'September',
        'October',
        'November',
        'December',
    ],
    shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
    shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
    goToToday: 'Go to today',
    prevMonthAriaLabel: 'Go to previous month',
    nextMonthAriaLabel: 'Go to next month',
    prevYearAriaLabel: 'Go to previous year',
    nextYearAriaLabel: 'Go to next year',
    closeButtonAriaLabel: 'Close date picker',
};
const controlClass = mergeStyleSets({
    control: {
        margin: '0 0 15px 0',
        width: ''

    },
});
export default class CreateRoute extends React.Component<IRouteProps, IRouteState> {
    public contactflag: any;
    public constructor(props: IRouteProps) {
        super(props);
        this.state = {
            planneddate: null,
            dealername: null,
            contactnumber: null,
            contactnumbererrormsg: "",
            remarks: "",
            plannedvisittime: "",
            location: "",
            assignto: null,
            district: null,
            dealeroption: [],
            locationoption: [],
            assigntooption: [],
            districtoption: []

        };
        this.dealerChanged = this.dealerChanged.bind(this);
        this._oncontactnumberchange = this._oncontactnumberchange.bind(this);
        this.locationChange = this.locationChange.bind(this);
        this.assigntoChange = this.assigntoChange.bind(this);
        this.districtChange = this.districtChange.bind(this);
    }
    private _onCancel = () => {
        this.props.onClose();
        this.setState({
            planneddate: null,
            dealername: null,
            contactnumber: null,
            contactnumbererrormsg: "",
            remarks: "",
            plannedvisittime: "",
            location: null,
            assignto: null,
            district: null
        })
    }
    public dealerarray = [];
    public async componentDidMount() {

        let locationarray = [];
        let assigntoarray = [];
        let districtarray = [];
        let dealerarray = [];
        // const dealeritems: any[] = await sp.web.lists.getByTitle("Dealer List").items.select("Title,ID").getAll();
        // console.log("dealer" + dealeritems);
        // for (let i = 0; i < dealeritems.length; i++) {

        //     let data = {
        //         key: dealeritems[i].Id,
        //         text: dealeritems[i].Title
        //     };

        //     dealerarray.push(data);
        // }
        // this.setState({
        //     dealeroption: dealerarray
        // });
        const locationitems: any[] = await sp.web.lists.getByTitle("Location").items.select("Title,ID").getAll();
        console.log("location" + locationitems);
        for (let i = 0; i < locationitems.length; i++) {

            let data = {
                key: locationitems[i].Id,
                text: locationitems[i].Title
            };

            locationarray.push(data);
        }
        this.setState({
            locationoption: locationarray
        });
        // const salesuseritems: any[] = await sp.web.lists.getByTitle("Users").items.select("Title,ID").getAll();
        // console.log("salesusers" + salesuseritems);
        // for (let i = 0; i < salesuseritems.length; i++) {

        //     let data = {
        //         key: salesuseritems[i].Id,
        //         text: salesuseritems[i].Title
        //     };

        //     assigntoarray.push(data);
        // }
        // this.setState({
        //     assigntooption: assigntoarray
        // });
        const districtitems: any[] = await sp.web.lists.getByTitle("Districts").items.select("Title,ID").getAll();
        console.log("district" + districtitems);
        for (let i = 0; i < districtitems.length; i++) {

            let data = {
                key: districtitems[i].Id,
                text: districtitems[i].Title
            };

            districtarray.push(data);
        }
        this.setState({
            districtoption: districtarray
        });


    }

    public _onplanneddateChange = (date?: Date): void => {
        this.setState({ planneddate: date });

        console.log(this.state.planneddate);
    }
    public async dealerChanged(option: { key: any; }) {
        //console.log(option.key);
        let locationarray = [];
        this.setState({ dealername: option.key });
        const locationitems: any[] = await sp.web.lists.getByTitle("Location").items.select("Title,ID").filter(" DistrictId eq " + option.key).get();
        console.log("location" + locationitems);
        for (let i = 0; i < locationitems.length; i++) {

            let data = {
                key: locationitems[i].Id,
                text: locationitems[i].Title
            };

            locationarray.push(data);
        }
        this.setState({
            locationoption: locationarray
        });
    }
    public locationChange(option: { key: any; }) {
        //console.log(option.key);
        this.setState({ location: option.key });
        console.log(this.state.location);
    }
    public async districtChange(option: { key: any; }) {
        let dealerarray = [];
        let assigntoarray = [];
        this.setState({ district: option.key });
        const dealeritems: any[] = await sp.web.lists.getByTitle("Dealer List").items.select("Title,ID").filter(" DistrictId eq " + option.key).get();

        console.log("dealer" + dealeritems);
        for (let i = 0; i < dealeritems.length; i++) {

            let data = {
                key: dealeritems[i].Id,
                text: dealeritems[i].Title
            };

            dealerarray.push(data);
        }
        this.setState({
            dealeroption: dealerarray
        });
        const salesuseritems: any[] = await sp.web.lists.getByTitle("Users").items.select("Title,ID").filter(" DistrictId eq " + option.key).get();
        console.log("salesusers" + salesuseritems);
        for (let i = 0; i < salesuseritems.length; i++) {

            let data = {
                key: salesuseritems[i].Id,
                text: salesuseritems[i].Title
            };

            assigntoarray.push(data);
        }
        this.setState({
            assigntooption: assigntoarray
        });
    }
    public _oncontactnumberchange = (ev: React.FormEvent<HTMLInputElement>, mob?: any) => {
        this.setState({ contactnumber: mob });
        let mnum = /^(\+\d{1,3}[- ]?)?\d{10}$/;
        let mnum2 = /^(\+\d{1,3}[- ]?)?\d{11}$/;
        //let mnum = /^(\+\d{1,3}[- ]?)$/;
        if (mob.match(mnum) || mob.match(mnum2) || mob == null) {
            this.setState({ contactnumbererrormsg: '' });
            this.contactflag = 1;

        }
        else {
            this.setState({ contactnumbererrormsg: 'Please enter a valid mobile number' });
            this.contactflag = 0;
        }
    }
    public onplannedvisittimechange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {

        //alert(newValue);
        this.setState({ plannedvisittime: newValue });


    }
    public remarkschange = (ev: React.FormEvent<HTMLInputElement>, remarks?: any) => {

        this.setState({ remarks: remarks });

    }
    public assigntoChange(option: { key: any }) {
        this.setState({ assignto: option.key });
    }
    public update = async () => {

        let siteUrl = "https://mrbutlers.sharepoint.com/sites/SalesOfficerApplication";
        let web = Web(siteUrl);
        let planneddate = moment(this.state.planneddate, 'DD/MM/YYYY').format("DD MMM YYYY");

        let conf = confirm("Do you want to submit?");
        if (conf == true) {

            sp.web.lists.getByTitle("Route List").items.add({

                Title: this.state.plannedvisittime,
                PlannedDate: planneddate,
                DistrictId: this.state.district,
                DealerNameId: this.state.dealername,
                ContactNumber: this.state.contactnumber,
                LocationId: this.state.location,
                AssignToId: this.state.assignto,
                Remarks: this.state.remarks


            }).then(i => {
                this._onCancel();
            })
        }

    }
    public render(): React.ReactElement<IRouteProps> {
        const { firstDayOfWeek } = this.state;
        let { isOpen } = this.props;
        return (

            <Panel isOpen={isOpen} type={PanelType.custom}
                customWidth={'800px'} onDismiss={this._onCancel}>
                <h3>Create Route</h3>

                <Label>Planned Date</Label>

                <DatePicker //style={{ width: '1000px' }}
                    //className={controlClass.control}
                    firstDayOfWeek={firstDayOfWeek}
                    strings={DayPickerStrings}
                    value={this.state.planneddate}
                    onSelectDate={this._onplanneddateChange}
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    isRequired={true}
                />
                <p><Label >Select District</Label>  <Dropdown id="dept"
                    placeholder="Select an option"
                    selectedKey={this.state.district}
                    options={this.state.districtoption}
                    //onChanged={this.dChanged}
                    onChanged={this.districtChange}
                    required={true}
                /></p>
                <Label >Dealer Name</Label>  <Dropdown id="dept"
                    placeholder="Select an option"
                    selectedKey={this.state.dealername}
                    options={this.state.dealeroption}
                    onChanged={this.dealerChanged}
                    required={true}
                //onChange={this.deptChanged}
                />
                <p><Label >Contact Number </Label>
                    < TextField value={this.state.contactnumber} onChange={this._oncontactnumberchange} errorMessage={this.state.contactnumbererrormsg} required={true}   ></TextField></p>
                <TextField
                    id="time"
                    label="Planned Visit Time"
                    type="time"
                    //defaultValue="07:30"
                    value={this.state.plannedvisittime}
                    onChange={this.onplannedvisittimechange}
                    required={true}
                />
                <p><Label >Location</Label>  <Dropdown id="dept"
                    placeholder="Select an option"
                    selectedKey={this.state.location}
                    options={this.state.locationoption}
                    //onChanged={this.dChanged}
                    onChanged={this.locationChange}
                /></p>
                <p><Label >Assign To</Label>
                    <Dropdown id="assign"
                        placeholder="Select an option"
                        selectedKey={this.state.assignto}
                        options={this.state.assigntooption}
                        //onChanged={this.dChanged}
                        onChanged={this.assigntoChange}
                    /></p>
                <p><Label >Remarks</Label>
                    < TextField value={this.state.remarks} onChange={this.remarkschange} multiline  ></TextField></p>

                <DialogFooter>
                    <PrimaryButton text="Save" onClick={this.update} />
                    <PrimaryButton text="Cancel" onClick={this._onCancel} />
                </DialogFooter>
            </Panel>

        );
    }



}