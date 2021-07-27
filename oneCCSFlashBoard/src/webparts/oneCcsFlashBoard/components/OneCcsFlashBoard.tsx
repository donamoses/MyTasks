import * as React from 'react';
import styles from './OneCcsFlashBoard.module.scss';
import { IOneCcsFlashBoardProps } from './IOneCcsFlashBoardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Marquee from "react-simple-marquee";
import { Items, sp } from "@pnp/sp/presets/all";
import * as ReactDOM from 'react-dom';

export interface IOneCcsFlashBoardState {
  ourFocus: string;
  description: any[];
  carouselElements: any[];

}
export interface ourFocusItems {
  Category: any;


}

var firstItem: ourFocusItems[] = [];
let final;
export default class OneCcsFlashBoard extends React.Component<IOneCcsFlashBoardProps, IOneCcsFlashBoardState, {}> {
  constructor(props: IOneCcsFlashBoardProps) {
    super(props);
    this.state = {
      ourFocus: "",
      description: [],
      carouselElements: [],
    };
  }

  public async componentDidMount() {
    this._ApplicationCategory();
    console.log(this.props.backGround);

  }

  private _ApplicationCategory = () => {
    let Vision;

    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.listName).items.get().then(results => {

      for (let i = 0; i < results.length; i++) {
        firstItem.push({
          Category: results[i].Category + ": " + results[i].Title,
        });
        console.log(firstItem);
      }
      for (let i = 0; i <= firstItem.length; i++) {
        let temp = firstItem[i].Category;
        firstItem[i].Category = firstItem[i + 1].Category;
        firstItem[i + 1].Category = temp;
        final = temp + this.props.seperator + firstItem[i].Category;
        console.log(final);
      }


      this.setState({
        ourFocus: final,
        description: firstItem,
      });
      console.log(final);
    });

  }


  public render(): React.ReactElement<IOneCcsFlashBoardProps> {

    return (
      <div>
        <Marquee>
          <div>{this.state.ourFocus}</div>
        </Marquee>
      </div>
    );
  }

}
