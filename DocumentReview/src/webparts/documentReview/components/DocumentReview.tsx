import * as React from 'react';
import styles from './DocumentReview.module.scss';
import { IDocumentReviewProps } from './IDocumentReviewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, DatePicker, DefaultButton, Dropdown, FontWeights, getTheme, IconButton, IDropdownOption, IIconProps, Label, mergeStyleSets, TextField } from 'office-ui-fabric-react';
import ReactHtmlParser, { processNodes, convertNodeToElement, htmlparser2 } from 'react-html-parser';
// import Moment from 'react-moment';
const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',
    height: '700px'
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      // borderTop: `10px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      padding: '12px 12px 14px 24px',

    },
  ],
  body: {
    width: '750px',
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    overflowY: 'hidden',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
});
const cancelIcon: IIconProps = { iconName: 'Cancel' };
const iconButtonStyles = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px',
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};
export interface IDocumentReviewState {
  // currentuser: any;
  // verifierId: any;
  // Reviewer: any;
  // approver: any;
  requestor:any;
  LinkToDoc: any;
  requestorComments:any;
  dueDate:any;
  // dcc: any;
  DCCComments:any;
  hideproject: boolean;
}
export default class DocumentReview extends React.Component<IDocumentReviewProps,IDocumentReviewState, {}> {
  public constructor(props: IDocumentReviewProps) {
    super(props);
    this.state = {
      // currentuser: "",
      // verifierId: "",
      // Reviewer: "",
      // approver: "",
      requestor:"",
      LinkToDoc: "",
      requestorComments:"",
      dueDate:"",
      // dcc: "",
      DCCComments:"",
       hideproject: true
    };
  }
  public async componentDidMount() {
      console.log(this.props.project);
    if (this.props.project) {
      this.setState({ hideproject: false });
    }
  }
  public render(): React.ReactElement<IDocumentReviewProps> {
    const Status: IDropdownOption[] = [

      { key: 'Reviewed', text: 'Reviewed' },
      { key: 'Cancelled', text: 'Cancelled' },
     
    ];
    return (
      <div className={ styles.documentReview }>
          <div className={contentStyles.header}>
          <span className={styles.title}>Review form of NOT/SHML/INT-PRC/AM-00009</span>
          <IconButton
            iconProps={cancelIcon}
            ariaLabel="Close popup modal"
            // onClick={this._closeModal}
            styles={iconButtonStyles}
          />
        </div>
        <div >
          <Label style={{ color: "red" }}>* fields are mandatory </Label>
          <Label >Document :  <a href={this.state.LinkToDoc}>NOT/SHML/INT-PRC/AM-00009 Migration Policy</a></Label>
          <table>
            <tr>
              <td><Label >Orginator : SUNIL JOHN </Label></td>
              <td><Label >Due Date : 21 JUL 2021 </Label></td>
              <td><Label >Revision : 0 </Label></td>
              <td hidden ={this.state.hideproject}><Label>Revision Level : ABT </Label></td>
            </tr>
          </table>
          <Label>Requestor : SUBHA RAVEENDRAN </Label> 
          <Label> Requestor Comment:</Label><div className={styles.commentdiv}>{ReactHtmlParser(this.state.requestorComments)}</div>
          <div hidden={this.state.hideproject}>
          <Label>DCC : SUBHA RAVEENDRAN </Label> 
          <Label>DCC Comment:</Label><div className={styles.commentdiv}>{ReactHtmlParser(this.state.DCCComments)}</div>
          <Checkbox label="Approve in same revision ? " boxSide="end" />
          </div>
          </div>
          <div style={{ marginTop: '30px' }}>
          <Dropdown 
          placeholder="Select Status" 
          label="Status"
          style={{ marginBottom: '10px', backgroundColor: "white" }}
          options={Status}
          // onChanged={this.ChangeId}
          // selectedKey={this.state.Status ? this.state.Status.key : undefined}
          required />

        <TextField label="Comments" id="Comments" multiline autoAdjustHeight />
        <DefaultButton id="b2" style={{ marginTop: '20px', float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }}>Submit</DefaultButton >
          <DefaultButton id="b2" style={{ marginTop: '20px', float: "right", marginRight: "10px", borderRadius: "10px", border: "1px solid gray" }}>Save</DefaultButton >
          <br />
        </div>
      </div>
    );
  }
}
