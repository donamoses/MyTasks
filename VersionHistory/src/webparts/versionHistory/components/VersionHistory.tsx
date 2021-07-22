import * as React from 'react';
import styles from './VersionHistory.module.scss';
import { IVersionHistoryProps } from './IVersionHistoryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { sp } from '@pnp/sp';
import "@pnp/sp/sites";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
export interface IVersionHistoryState {
  siteurl:any;
}
export default class VersionHistory extends React.Component<IVersionHistoryProps,IVersionHistoryState, {}> {
  public constructor(props: IVersionHistoryProps) {
    super(props);
    this.state = {
      siteurl:""
    };
  }
  public async componentDidMount() {
    const rootwebData = await sp.site.rootWeb();
    console.log(rootwebData);
    var webValue = rootwebData.ResourcePath.DecodedUrl;
    //alert(webValue);
   
    this.setState({
      siteurl: webValue
    });
   
}

  public render(): React.ReactElement<IVersionHistoryProps> {
    return (
      <div className={ styles.versionHistory }>
       <IFrameDialog
                        url={this.state.siteurl + "/_layouts/15/Versions.aspx?list=%7Bda53146b-3f5c-4321-926e-c3c2adbff323%7D&ID=1&IsDlg=0"}
                        title="Version History"
                        hidden={false}
                        // onDismiss={this.onCancel}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: "700px !important", width: "600px !important", height: "800px !important" } }
                        }}
                        dialogContentProps={{
                            type: DialogType.close,
                            showCloseButton: true
                        }}
                        width={'800px'}
                        height={'500px'}
                    />
      </div>
    );
  }
}
