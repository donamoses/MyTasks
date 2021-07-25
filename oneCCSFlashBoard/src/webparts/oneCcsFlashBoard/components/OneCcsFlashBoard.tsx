import * as React from 'react';
import styles from './OneCcsFlashBoard.module.scss';
import { IOneCcsFlashBoardProps } from './IOneCcsFlashBoardProps';
import { escape } from '@microsoft/sp-lodash-subset';
//import Marquee from "react-simple-marquee";
import { sp } from "@pnp/sp/presets/all";
import { hasHorizontalOverflow, VerticalDivider } from 'office-ui-fabric-react';
import Marquee from 'react-simple-marquee';
export interface IOneCcsFlashBoardState {
  ourFocus: string;
  description: any[];

}

export default class OneCcsFlashBoard extends React.Component<IOneCcsFlashBoardProps, IOneCcsFlashBoardState, {}> {
  constructor(props: IOneCcsFlashBoardProps) {
    super(props);
    this.state = {
      ourFocus: "",
      description: [],
    };
  }

  public async componentDidMount() {
    this._ApplicationCategory();

  }
  private _ApplicationCategory = () => {
    let myarray = "";
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.listName).items.select("Title,Category,Description0").get().then(results => {
      // let marquee = document.createElement("DIV");
      for (let i = 0; i < results.length; i++) {
        myarray += results[i].Description0;
      }
      // let outputText = marquee.innerText;
      //console.log((marquee.innerText).concat(this.props.seperator));
      //console.log(outputText);
      return this.setState({
        ourFocus: myarray,
        description: results,
      });
    });

  }
  private createMarkup() {
    return { __html: this.state.ourFocus };
  }
  public render(): React.ReactElement<IOneCcsFlashBoardProps> {
    return (
      <div className={styles.oneCcsFlashBoard}>
        <div className={styles.container}>
          {this.state.description.map((items) => {
            <div style={{ display: " Marquee" }}> hi </div>
            if (this.props.horizontal) {
              return (<Marquee
                speed={1} // Speed of the marquee (Optional)
                style={{
                  height: 50 // Your own styling (Optional)
                }}

              >
                <div dangerouslySetInnerHTML={this.createMarkup()}></div>

              </Marquee>)

            }
            else if (this.props.vertical) {
              return (
                <Marquee
                  speed={1} // Speed of the marquee (Optional)
                  style={{
                    height: 50 // Your own styling (Optional)
                  }}
                  direction={'bottom-top'}
                >
                  <h4>{items.Title}</h4>
                </Marquee>
              )

            }

          })}
        </div>
      </div >

    );
  }
}
