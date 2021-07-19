import * as React from 'react';
import styles from './RevisionHistory.module.scss';
import { IRevisionHistoryProps } from './IRevisionHistoryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { VerticalTimeline, VerticalTimelineElement } from 'react-vertical-timeline-component';
import 'react-vertical-timeline-component/style.min.css';
//import { WorkIcon, StarIcon } from 'VerticalTimelineElement';
import { Item, sp } from "@pnp/sp/presets/all";
import * as moment from 'moment';
import { Icon } from 'office-ui-fabric-react';
import * as strings from "RevisionHistoryWebPartStrings";

import SchoolIcon from '@material-ui/icons/School';
import WorkIcon from '@material-ui/icons/Work';
import StarIcon from '@material-ui/icons/Star';
import LoyaltyIcon from '@material-ui/icons/Loyalty';


//import "src/VerticalTimeline.css";
//import "src/VerticalTimelineElement.css";
export interface IRevisionHistoryState {
  docRepositoryItems: any[];
  query: any;
  title: string;
  documentName: string;
  createdDate: any;
  currentUser: any;
  wFlowDetailItems: any[];
  wFlowStatus: string;
  Revision: String;
  Verifier: string;
  Approver: string;
}

const workIcon = () => <Icon iconName="TextDocumentSettings" />;
export default class RevisionHistory extends React.Component<IRevisionHistoryProps, IRevisionHistoryState, {}> {
  constructor(props: IRevisionHistoryProps) {
    super(props);
    this.state = {
      docRepositoryItems: [],
      query: "",
      title: "",
      documentName: "",
      createdDate: null,
      currentUser: "",
      wFlowDetailItems: [],
      wFlowStatus: "",
      Revision: "",
      Verifier: "",
      Approver: "",
    };

  }


  public async componentDidMount() {
    const rootwebData = await sp.site.rootWeb();
    console.log(rootwebData);
    var webValue = rootwebData.ResourcePath.DecodedUrl;
    //alert(webValue);
    this.setState({
      // siteurl: webValue
    });
    //url query getting id
    let params = new URLSearchParams(window.location.search);
    let id = params.get('DRID');
    console.log(id);
    if (id != undefined && id != null) {
      this.setState({
        query: Number(id),
      });

      //getting current user
      sp.web.currentUser.get().then(r => {
        console.log(r);
        this.setState({
          currentUser: r.Title,
        });
      });
    }
    this.loadDocProfile();

  }
  private loadDocProfile = async () => {
    //getting list DocProfile u
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.listName).items.select("Title,Created,DocumentName,DocumentResponsible/Title,DocumentResponsible/ID,ID,Revision,Created,Author/ID,Author/Title,WFStatus,Approver/ID,Approver/Title,Verifier/ID,Verifier/Title").expand("DocumentResponsible,Author,Approver,Verifier").filter("ID eq " + this.state.query + " ").get().then(docProfileItems => {

      this.setState({
        docRepositoryItems: docProfileItems,
        createdDate: moment.utc(docProfileItems[0].Created).format('DD-MMM-YYYY'),
        documentName: docProfileItems[0].DocumentName,
        wFlowStatus: docProfileItems[0].wFlowStatus,
        Revision: docProfileItems[0].Revision,
        Verifier: docProfileItems[0].Verifier.Title,
        Approver: docProfileItems[0].Approver.Title,
      });
      console.log(this.state.docRepositoryItems);
    });
    //getting list from WF Details
    sp.web.getList("/sites/DMS/Lists/DMSWFDetails").items.select("Action,PrevStatus,Approver/ID,Approver/Title,Requestor/ID,Requestor/Title,RequestDate,ResponseDate").expand("Approver,Requestor").filter("DocumentRepositoryIDId eq '" + this.state.query + "'").get().then(wfDetailsItems => {

      this.setState({
        wFlowDetailItems: wfDetailsItems,
      });
      console.log(this.state.wFlowDetailItems);
    });
  }
  private draftData = (item: any) => {

    return (<VerticalTimelineElement
      className="vertical-timeline-element--work"
      contentStyle={{ background: 'rgb(155 216 235 / 65%);', color: 'rgb(19 18 18)' }}
      contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
      date={moment.utc(item.Created).format('DD-MMM-YYYY')}
      iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
      icon={<WorkIcon />}
    >
      <div style={{ fontSize: "20px", color: "rgb(220,20,60)" }}> Document Created </div>
      <br></br>
      <h4 className="vertical-timeline-element-subtitle">  {item.DocumentName}</h4>
      <p>

        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Created By:{item.Author.Title}
        </div>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>

        </div>
        <br></br>
      </p>
    </VerticalTimelineElement>);
  }
  private workFlowStarted = (wFDitem: any) => {
    return (<VerticalTimelineElement
      className="vertical-timeline-element--work"
      contentStyle={{ background: 'rgb(185 237 137)', color: 'rgb(19 18 18)' }}
      contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
      date={moment.utc(wFDitem.RequestDate).format('DD-MMM-YYYY')}
      iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
      icon={<WorkIcon />}
    >
      <div style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>WorkFlow Started </div>
      <br></br>
      <h4 className="vertical-timeline-element-subtitle">  {this.state.documentName}</h4>
      <p>
        <div> Requestor : {wFDitem.Requestor.Title}</div>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Verifier :{this.state.Verifier} <td></td>   <div style={{ margin: "0px 0px 0px 22px" }}>Approver : {this.state.Approver}</div>

        </div>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Version :{this.state.Revision}
        </div>
        <br></br>
      </p>
    </VerticalTimelineElement>);
  }
  private Verified = (ver: any) => {
    return (<VerticalTimelineElement
      className="vertical-timeline-element--work"
      contentStyle={{ background: 'rgb(213 202 231)', color: 'rgb(19 18 18)' }}
      contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
      date={moment.utc(ver.ResponseDate).format('DD-MMM-YYYY')}
      iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
      icon={<SchoolIcon />}
    >
      <div style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>Verified </div>
      <br></br>
      <h4 className="vertical-timeline-element-subtitle">  {this.state.documentName}</h4>
      <p>
        Requestor: {ver.Requestor.Title}
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Verified By : {this.state.Verifier}
        </div>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Version:{this.state.Revision}
        </div>
        <br></br>
      </p>
    </VerticalTimelineElement>);
  }
  private Approved = (apr: any) => {
    return (<VerticalTimelineElement
      className="vertical-timeline-element--work"
      contentStyle={{ background: 'rgb(233 205 126)', color: 'rgb(19 18 18)' }}
      contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
      date={moment.utc(apr.ResponseDate).format('DD-MMM-YYYY')}
      iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
      icon={<LoyaltyIcon />}
    >
      <div style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>Approved</div>
      <br></br>
      <h4 className="vertical-timeline-element-subtitle">  {this.state.documentName}</h4>
      <p>
        <div>{apr.Title}</div>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Approved By :  {this.state.Approver}
        </div>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Version : {this.state.Revision}
        </div>
        <br></br>
      </p>
    </VerticalTimelineElement>);
  }
  private Published = (pub: any) => {
    return (<VerticalTimelineElement
      className="vertical-timeline-element--work"
      contentStyle={{ background: 'rgb(233 157 127)', color: 'rgb(19 18 18)' }}
      contentArrowStyle={{ borderRight: '7px solid  rgb(19 18 18)' }}
      date={moment.utc(pub.ResponseDate).format('DD-MMM-YYYY')}
      iconStyle={{ background: 'rgb(0, 120, 212)', color: '#fff' }}
      icon={<LoyaltyIcon />}
    >
      <div style={{ fontSize: "20px", color: "rgb(220,20,60)" }}>Published </div>
      <br></br>
      <h4 className="vertical-timeline-element-subtitle">  {this.state.documentName}</h4>
      <p>
        Verified By : {this.state.Verifier}
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Approved By :{this.state.Approver}
        </div>
        <div style={{ display: 'flex', margin: "0px 0px 0px 0px" }}>
          Version : {this.state.Revision}
        </div>
        <br></br>
      </p>
    </VerticalTimelineElement >);
  }
  public render(): React.ReactElement<IRevisionHistoryProps> {

    return (

      <div style={{ width: "110%" }}>
        <h1>Flow History</h1>
        <div>
          <VerticalTimeline>
            {
              this.state.docRepositoryItems.map((item, key) => {
                if (item.WFStatus == "Draft" || item.WFStatus == "Under verification" || item.WFStatus == "Verified" || item.WFStatus == "Published") {
                  return (this.draftData(item));
                }
              })
            }
            {
              this.state.wFlowDetailItems.map((WFDitem, key) => {
                if (WFDitem.PrevStatus == "Under verification") {
                  return (this.workFlowStarted(WFDitem));
                }

              })
            }
            {
              this.state.wFlowDetailItems.map((ver, key) => {

                if (ver.PrevStatus == "Under verification" || ver.Action == "VER") {
                  return (this.Verified(ver));
                }
              })
            }
            {
              this.state.wFlowDetailItems.map((apr, key) => {

                if (apr.PrevStatus == "Published" || apr.Action == "APR") {
                  return (this.Approved(apr));
                }
              })
            }
            {
              this.state.wFlowDetailItems.map((pub, key) => {

                if (pub.PrevStatus == "Published" || pub.Action == "APR") {
                  return (this.Published(pub));
                }
              })
            }
            <VerticalTimelineElement
              iconStyle={{ background: ' #f1083e', color: '#fff' }}
            // icon={<StarIcon />}
            />
          </VerticalTimeline>
        </div>
      </div>
    );
  }

}


