
import { DatePicker, IconButton, NormalPeoplePicker, Panel, PanelType } from '@fluentui/react';
import * as React from 'react';
import { ICreateEmployeeProps } from './ICreateEmployeeProps';
import { sp } from "@pnp/sp/presets/all";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Checkbox, ChoiceGroup, DialogFooter, Dropdown, IChoiceGroupOption, IDropdownOption, Label, MessageBar, MessageBarType, Pivot, PivotItem, PrimaryButton, TextField } from 'office-ui-fabric-react';
import styles from './CreateEmployee.scss';
import SimpleReactValidator from 'simple-react-validator';
export interface ICreateEmployeeState {
    selectedDept: any;
    selectedDeptText: string;
    selectedDeptKey: number;
    selectedDesigKey: number;
    selectedDesigText: string;
    designationDetails: any[];
    selectedQualiKey: number;
    selectedQualiText: string;
    qualificationDetails: any[];
    selectedLocaKey: number;
    selectedLocaText: string;
    locationDetails: any[];
    selectedROKey: number;
    selectedROText: string;
    reportingOfficerDetails: any[];
    selectedNatKey: number;
    selectedNatText: string;
    nationalityDetails: any[];
    birthdate: any;
    joindate: any;
    firstname: string;
    selectedgender: any;
    lastName: any;
    mobile: any;
    email: string;
    alterEmail: string;
    alterPhone: any;
    workPhone: any;
    perAddress: any;
    temAddress: any;
    country: any;
    city: any;
    state: any;
    pin: any;
    idProof: any;
    passport: any;
    license: any;
    address: any;
    userNameId: any;
    repotingOfficerKey: any;
    userEmail: any;
    empid: any;
    noticePeriod: any;
    EmployeeAdded: any;
}
var options: IDropdownOption[];
var desigOptions: IDropdownOption[];
var nationalityoptions: IDropdownOption[];
var locationOptions: IDropdownOption[];
var qualifiOptions: IDropdownOption[];
var repofficerOptions: IDropdownOption[];
export default class CreateEmployee extends React.Component<ICreateEmployeeProps, ICreateEmployeeState, any> {
    private validator: SimpleReactValidator;

    constructor(props: ICreateEmployeeProps) {
        super(props);
        this.state = ({
            selectedDept: "",
            selectedDeptText: "",
            selectedDeptKey: null,
            selectedDesigKey: null,
            selectedDesigText: "",
            designationDetails: [],
            selectedQualiKey: null,
            selectedQualiText: "",
            qualificationDetails: [],
            selectedLocaKey: null,
            selectedLocaText: "",
            locationDetails: [],
            selectedROKey: null,
            selectedROText: "",
            reportingOfficerDetails: [],
            selectedNatKey: null,
            selectedNatText: "",
            nationalityDetails: [],
            birthdate: "",
            joindate: "",
            firstname: "",
            selectedgender: "",
            lastName: "",
            mobile: null,
            email: "",
            alterEmail: "",
            alterPhone: "",
            workPhone: "",
            perAddress: "",
            temAddress: "",
            country: "",
            city: "",
            state: "",
            pin: "",
            passport: "",
            idProof: "",
            license: "",
            address: "",
            userNameId: "",
            repotingOfficerKey: "",
            userEmail: "",
            empid: "",
            noticePeriod: "",
            EmployeeAdded: 'none',

        });
        this._onCancel = this._onCancel.bind(this);
        this._handleDepartment = this._handleDepartment.bind(this);
        this._handleLocation = this._handleLocation.bind(this);
        this._handleNationality = this._handleNationality.bind(this);
        this._handleQualification = this._handleQualification.bind(this);
        this._handleRepofficer = this._handleRepofficer.bind(this);
        this._handleDesignation = this._handleDesignation.bind(this);
        this._genderonChange = this._genderonChange.bind(this);
        this._onSaveEmployee = this._onSaveEmployee.bind(this);
        this._onEmailChange = this._onEmailChange.bind(this);
        this._onALterEmailChange = this._onALterEmailChange.bind(this);
        this._onAddressChange = this._onAddressChange.bind(this);
        this._onCityChange = this._onCityChange.bind(this);
        this._onCountryChange = this._onCountryChange.bind(this);
        this._onEmailChange = this._onEmailChange.bind(this);
        this._onLastNameChange = this._onLastNameChange.bind(this);
        this._onFirstNameChange = this._onFirstNameChange.bind(this);

    }

    public async componentDidMount() {
        this._department();
        this._designation();
        this._nationality();
        this._location();
        this._qualification();
        this._repoOfficer();
    }
    public async componentWillMount() {
        this.validator = new SimpleReactValidator({
            messages: {
                required: "Please enter mandatory fields"
            }
        });

    }
    private _onCancel = () => {
        this.props.onClose();
        this.setState({
            selectedDept: "",
            birthdate: "",
            joindate: "",
            firstname: "",
            selectedgender: "",
            lastName: "",
            mobile: null,
            email: "",
            alterEmail: "",
            alterPhone: "",
            workPhone: "",
            perAddress: "",
            temAddress: "",
            country: "",
            city: "",
            state: "",
            pin: "",
            passport: "",
            idProof: "",
            license: "",
            address: "",
            userNameId: "",
            repotingOfficerKey: "",
            userEmail: "",
            empid: "",
            noticePeriod: "",

        });
    }
    private _department = () => {
        sp.web.getList("sites/CCSHR/Lists/Department").items.get().then(dept => {
            console.log(dept);
            options = [];
            for (let k in dept) {
                options.push(
                    { key: dept[k].ID, text: dept[k].Title }
                );
            }
            return this.setState({
                selectedDept: dept,

            });
        });
    }
    private _designation = () => {

        sp.web.getList("sites/CCSHR/Lists/Designation").items.get().then(desig => {
            console.log(desig);
            desigOptions = [];
            for (let k in desig) {
                desigOptions.push(
                    { key: desig[k].ID, text: desig[k].Title }
                );
            }
            return this.setState({
                designationDetails: desig,

            });
        });
    }
    private _location = () => {

        sp.web.getList("sites/CCSHR/Lists/Location").items.get().then(loc => {
            console.log(loc);
            locationOptions = [];
            for (let k in loc) {
                locationOptions.push(
                    { key: loc[k].ID, text: loc[k].Title }
                );
            }
            return this.setState({
                locationDetails: loc,

            });
        });
    }
    private _qualification = () => {

        sp.web.getList("sites/CCSHR/Lists/Qualification").items.get().then(quali => {
            console.log(quali);
            qualifiOptions = [];
            for (let k in quali) {
                qualifiOptions.push(
                    { key: quali[k].ID, text: quali[k].Title }
                );
            }
            return this.setState({
                qualificationDetails: quali,

            });
        });
    }
    private _nationality = () => {
        sp.web.getList("sites/CCSHR/Lists/Nationality").items.get().then(nat => {
            console.log(nat);
            nationalityoptions = [];
            for (let k in nat) {
                nationalityoptions.push(
                    { key: nat[k].ID, text: nat[k].Title }
                );
            }
            return this.setState({
                nationalityDetails: nat,

            });
        });
    }
    private _repoOfficer = () => {
        sp.web.getList("sites/CCSHR/Lists/Employees").items.get().then(ro => {
            console.log(ro);
            repofficerOptions = [];
            for (let k in ro) {
                repofficerOptions.push(
                    { key: ro[k].ID, text: ro[k].FullName }
                );
            }
            return this.setState({
                reportingOfficerDetails: ro,

            });
        });
    }
    private _handleDepartment(selectedOption: { key: any; text: any; }) {
        console.log(selectedOption.text);
        this.setState({
            selectedDeptKey: selectedOption.key,
            selectedDeptText: selectedOption.text,
        });

    }
    private _handleDesignation(selectedOption: { key: any; text: any; }) {
        console.log(selectedOption.text);
        this.setState({
            selectedDesigKey: selectedOption.key,
            selectedDesigText: selectedOption.text,
        });

    }


    private _handleLocation(selectedOption: { key: any; text: any; }) {
        console.log(selectedOption.text);
        this.setState({
            selectedLocaKey: selectedOption.key,
            selectedLocaText: selectedOption.text,
        });

    }
    private _handleNationality(selectedOption: { key: any; text: any; }) {
        console.log(selectedOption.text);
        this.setState({
            selectedNatKey: selectedOption.key,
            selectedNatText: selectedOption.text,
        });

    }
    private _handleQualification(selectedOption: { key: any; text: any; }) {
        console.log(selectedOption.text);
        this.setState({
            selectedQualiKey: selectedOption.key,
            selectedQualiText: selectedOption.text,
        });

    }
    private _handleRepofficer(selectedOption: { key: any; text: any; }) {
        console.log(selectedOption.text);
        this.setState({
            selectedROKey: selectedOption.key,
            selectedROText: selectedOption.text,
        });

    }

    public _onbirthDatePickerChange = (date?: Date): void => {
        this.setState({ birthdate: date.toLocaleString() });
        console.log(date);
        var getcurrentyear = new Date();
        var currentyear = getcurrentyear.getFullYear();

        var selectedYear = date.getFullYear(); // selected year
        console.log(selectedYear);
        let agediff = currentyear - selectedYear;
        console.log(agediff);
        // this.setState({ age: agediff });
        // this.setState({ agehidden: false })

    }
    private _onjoinDatePickerChange = (date?: Date): void => {

        this.setState({ joindate: date.toLocaleString });
        console.log(this.state.joindate);

    }
    private _onFirstNameChange = (ev: React.FormEvent<HTMLInputElement>, newfname?: string) => {
        this.setState({ firstname: newfname });
    }
    private _onLastNameChange = (ev: React.FormEvent<HTMLInputElement>, newlname?: string) => {
        this.setState({ lastName: newlname || '' });
    }
    private _onWorkPhoneChange = (ev: React.FormEvent<HTMLInputElement>, WorkPhone?: string) => {
        this.setState({ workPhone: WorkPhone || '' });
    }
    private _onEmailChange = (ev: React.FormEvent<HTMLInputElement>, Email?: string) => {
        this.setState({ email: Email || '' });
    }
    private _onCountryChange = (ev: React.FormEvent<HTMLInputElement>, Country?: string) => {
        this.setState({ country: Country || '' });
    }
    private _onCityChange = (ev: React.FormEvent<HTMLInputElement>, City?: string) => {
        this.setState({ city: City || '' });
    }
    private _onStateChange = (ev: React.FormEvent<HTMLInputElement>, State?: string) => {
        this.setState({ state: State || '' });
    }
    private _onPINChange = (ev: React.FormEvent<HTMLInputElement>, PIN?: string) => {
        this.setState({ pin: PIN || '' });
    }
    private _onAlterPhoneChange = (ev: React.FormEvent<HTMLInputElement>, AlterPhone?: string) => {
        this.setState({ alterPhone: AlterPhone || '' });
    }
    private _onPAddressChange = (ev: React.FormEvent<HTMLInputElement>, Padd?: string) => {
        this.setState({ perAddress: Padd || '' });
    }
    private _onTAddressChange = (ev: React.FormEvent<HTMLInputElement>, Tadd?: string) => {
        this.setState({ temAddress: Tadd || '' });
    }
    private _onIdProofChange = (ev: React.FormEvent<HTMLInputElement>, Idp?: string) => {
        this.setState({ idProof: Idp || '' });
    }
    private _onLicenseChange = (ev: React.FormEvent<HTMLInputElement>, License?: string) => {
        this.setState({ license: License || '' });
    }
    private _onPassportChange = (ev: React.FormEvent<HTMLInputElement>, Passport?: string) => {
        this.setState({ passport: Passport || '' });
    }
    private _onMobileChange = (ev: React.FormEvent<HTMLInputElement>, Mobile?: string) => {
        this.setState({ mobile: Mobile || '' });
    }
    private _onALterEmailChange = (ev: React.FormEvent<HTMLInputElement>, AlterEmail?: string) => {
        this.setState({ alterEmail: AlterEmail || '' });
    }
    private _onAddressChange = (ev: React.FormEvent<HTMLInputElement>, Address?: string) => {
        this.setState({ address: Address || '' });
    }
    public _genderonChange = (ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): void => {
        console.log(option);
        this.setState({ selectedgender: option.key });
        console.log(this.state.selectedgender);
    }
    private _getUserNamePeoplePickerItems(approverItems: any[]) {
        if (approverItems.length != 0) {
            var userValue = [];
            userValue.push(approverItems[0]['text']);
            sp.web.ensureUser(approverItems[0]["text"]).then(userId => {
                this.setState({
                    userNameId: userId.data.Id,
                    userEmail: userId.data.Email,

                });
            });
        }
        else {
            this.setState({
                userNameId: "",
                userEmail: "",
            });
        }
        console.log('Items:', approverItems);
    }
    private _onAddressChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
        if (isChecked) {
            this.setState({ temAddress: this.state.perAddress });
        }
        else {
            this.setState({ temAddress: "" });

        }
    }

    private _onSaveEmployee = async () => {

        if (this.validator.fieldValid("firstname") && this.validator.fieldValid("lastName") && this.validator.fieldValid("joindate")
            && this.validator.fieldValid("selectedDeptKey") && this.validator.fieldValid("selectedDesigKey") && this.validator.fieldValid("email")) {
            sp.web.getList("sites/CCSHR/Lists/Employees").items.add(
                {
                    FirstName: this.state.firstname,
                    LastName: this.state.lastName,
                    DateOfbirth: this.state.birthdate,
                    DateOfJoining: this.state.joindate,
                    PermanentAddress: this.state.perAddress,
                    TemporaryAddress: this.state.temAddress,
                    Country: this.state.country,
                    State: this.state.state,
                    City: this.state.city,
                    Gender: this.state.selectedgender,
                    NationalityId: this.state.selectedNatKey,
                    DepartmentId: this.state.selectedDeptKey,
                    QualificationId: this.state.selectedQualiKey,
                    LocationId: this.state.selectedLocaKey,
                    DesignationId: this.state.selectedDesigKey,
                    IDProofDetails: this.state.idProof,
                    DrivingLicense: this.state.license,
                    PassportDetails: this.state.passport,
                    PIN: this.state.pin,
                    AlternativePhone: this.state.alterPhone,
                    AlternativeEmail: this.state.alterEmail,
                    EmailId: this.state.email,
                    MobileNo: this.state.mobile,
                    WorkPhone: this.state.workPhone,
                    UserNameId: this.state.userNameId,
                    ReportingOfficerId: this.state.selectedROKey,
                    Address: this.state.address,
                    NoticePeriod_x0028_inmonths_x002: this.state.noticePeriod,

                });
            this.validator.hideMessages();
            this.setState({ EmployeeAdded: '' });
            setTimeout(() => this.setState({ EmployeeAdded: 'none' }), 1000);

        }
        else {
            this.validator.showMessages();
            this.forceUpdate();
        }
        setTimeout(() => this.props.onClose(), 2000);

        //this._onCancel();
    }
    private _onNoticePeriod = (ev: React.FormEvent<HTMLInputElement>, NoticePeriod?: string) => {
        this.setState({ noticePeriod: NoticePeriod || '' });
    }

    public render(): React.ReactElement<ICreateEmployeeProps> {
        const gender: IChoiceGroupOption[] = [
            { key: 'Male', text: 'Male' },
            { key: 'Female', text: 'Female' }

        ];
        let { isOpen } = this.props;
        return (
            <Panel isOpen={isOpen} type={PanelType.custom}
                customWidth={'550px'} onDismiss={this._onCancel}>
                <IconButton iconProps={{ iconName: "PeopleAdd" }} title="Add Employee" label="Add Employee"></IconButton>

                <div>
                    <Pivot aria-label="Large Link Size Pivot Example">
                        <PivotItem headerText="Personal">

                            <TextField label="First Name" value={this.state.firstname} onChange={this._onFirstNameChange} required></TextField>
                            <div style={{ color: "#dc3545" }}>{this.validator.message("firstname", this.state.firstname, "required|alpha_space")}{" "}</div>
                            <TextField label="Last Name" value={this.state.lastName} onChange={this._onLastNameChange}></TextField>
                            <div style={{ color: "#dc3545" }}>{this.validator.message("lastName", this.state.lastName, "required|alpha_space")}{" "}</div>
                            <ChoiceGroup label="Gender" options={gender} onChange={this._genderonChange} value={this.state.selectedgender} required={true} />
                            <div style={{ color: "#dc3545" }}>{this.validator.message("firstname", this.state.firstname, "required")}{" "}</div>
                            <TextField label="Work Phone" onChange={this._onWorkPhoneChange} value={this.state.workPhone}></TextField>
                            <div style={{ color: "#dc3545" }}>{this.validator.message("workPhone", this.state.workPhone, "required|phone")}{" "}</div>
                            <TextField label="Address" onChange={this._onAddressChange} value={this.state.address} multiline></TextField>


                            <Label >Nationality</Label>
                            <Dropdown id="nat" required={true}
                                placeholder="Select an option"
                                selectedKey={this.state.selectedNatKey}
                                options={nationalityoptions}
                                onChanged={this._handleNationality}

                            />

                            <PeoplePicker
                                context={this.props.context}
                                titleText={"User Name"}
                                personSelectionLimit={1}
                                ensureUser={true}
                                showtooltip={true}
                                onChange={this._getUserNamePeoplePickerItems.bind(this)}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                resolveDelay={1000}
                            // defaultSelectedUsers={} 
                            />
                            <Label >Location</Label>
                            <Dropdown id="dept"
                                placeholder="Select an option"
                                selectedKey={this.state.selectedLocaKey}
                                options={locationOptions}
                                onChanged={this._handleLocation}

                            />
                            <Label >Qualification</Label>
                            <Dropdown id="dept"
                                placeholder="Select an option"
                                selectedKey={this.state.selectedQualiKey}
                                options={qualifiOptions}
                                onChanged={this._handleQualification}
                            />

                            <Label >Department</Label>
                            <Dropdown id="dept" required={true}
                                placeholder="Select an option"
                                selectedKey={this.state.selectedDeptKey}
                                options={options}
                                onChanged={(selectedOption) => this._handleDepartment(selectedOption)}

                            />
                            <div style={{ color: "#dc3545" }}>{this.validator.message("selectedDeptKey", this.state.selectedDeptKey, "required")},{" "}</div>
                            <Label >Designation</Label>
                            <Dropdown id="desig" required={true}
                                placeholder="Select an option"
                                selectedKey={this.state.selectedDesigKey}
                                options={desigOptions}
                                onChanged={this._handleDesignation}

                            />
                            <div style={{ color: "#dc3545" }}>{this.validator.message("selectedDesigKey", this.state.selectedDesigKey, "required")},{" "}</div>

                            <Label >Reporting Officer</Label>
                            <Dropdown id="dept"
                                placeholder="Select an option"
                                selectedKey={this.state.selectedROKey}
                                options={repofficerOptions}
                                onChanged={this._handleRepofficer}
                            />
                            <div style={{ color: "#dc3545" }}>{this.validator.message("selectedROKey", this.state.selectedROKey, "required")},{" "}</div>
                            <Label>Birth Date </Label>
                            <DatePicker style={{ width: '470px' }}
                                value={this.state.birthdate}
                                onSelectDate={this._onbirthDatePickerChange}
                                placeholder="Select a date..."
                                ariaLabel="Select a date"
                                maxDate={new Date()}

                            />
                            <div style={{ color: "#dc3545" }}>{this.validator.message("birthdate", this.state.birthdate, "required")},{" "}</div>
                            <Label>Joining Date</Label>

                            <DatePicker style={{ width: '470px' }}
                                value={this.state.joindate}
                                onSelectDate={this._onjoinDatePickerChange}
                                placeholder="Select a date..."
                                ariaLabel="Select a date"

                            />
                            <div style={{ color: "#dc3545" }}>{this.validator.message("joindate", this.state.joindate, "required")},{" "}</div>
                            <TextField label="Notice Period (in months)" onChange={this._onNoticePeriod} value={this.state.noticePeriod}></TextField>
                            <div style={{ color: "#dc3545" }}>{this.validator.message("noticePeriod", this.state.noticePeriod, "required|numeric|min:0,num")},{" "}</div>
                        </PivotItem>
                        <PivotItem headerText="Address">
                            <TextField label="Email" onChange={this._onEmailChange} value={this.state.email}></TextField>
                            <div style={{ color: "#dc3545" }}>{this.validator.message("email", this.state.email, "required|email")},{" "}</div>
                            <TextField label="Mobile" onChange={this._onMobileChange} value={this.state.mobile}></TextField>
                            <TextField label="Alternative Phone" onChange={this._onAlterPhoneChange} value={this.state.alterPhone}></TextField>
                            <TextField label="Country" onChange={this._onCountryChange} value={this.state.country}></TextField>
                            <TextField label="City" onChange={this._onCityChange} value={this.state.city}></TextField>
                            <TextField label="State" onChange={this._onStateChange} value={this.state.state}></TextField>
                            <TextField label="PIN" onChange={this._onPINChange} value={this.state.pin}></TextField>
                            <TextField label="Alternative Email" onChange={this._onALterEmailChange} value={this.state.alterEmail} ></TextField>
                            <div style={{ color: "#dc3545" }}>{this.validator.message("alterEmail", this.state.alterEmail, "required|email")},{" "}</div>
                            <TextField label="Permanent Address" onChange={this._onPAddressChange} value={this.state.perAddress} multiline></TextField>
                            <br></br>
                            <Checkbox label="Same as Permanent Address " onChange={this._onAddressChecked} />
                            <TextField label="Temporary Address" onChange={this._onTAddressChange} value={this.state.temAddress} multiline></TextField>

                        </PivotItem>
                        <PivotItem headerText="Additional">
                            <TextField label="ID Proof Details" onChange={this._onIdProofChange} value={this.state.idProof} multiline></TextField>
                            <TextField label="Driving License" onChange={this._onLicenseChange} value={this.state.license} multiline></TextField>
                            <TextField label="Passport Details" onChange={this._onPassportChange} value={this.state.passport} multiline></TextField>
                        </PivotItem>
                    </Pivot>
                </div>
                <DialogFooter>
                    <PrimaryButton text="Save" onClick={this._onSaveEmployee} />
                    <PrimaryButton text="Cancel" onClick={this._onCancel} />
                </DialogFooter>
                <div style={{ display: this.state.EmployeeAdded }}>
                    <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>  New Employee Addedd.</MessageBar>
                </div>
            </Panel >
        );
    }

}


