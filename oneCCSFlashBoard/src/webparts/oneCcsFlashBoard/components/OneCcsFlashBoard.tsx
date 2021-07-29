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
  vertical:string;
  carouselElements: any[];

}
export interface ourFocusItems {
  Category: any;
  verCatogary:any

}

var firstItem: ourFocusItems[] = [];
let final;
let verticalItem;
export default class OneCcsFlashBoard extends React.Component<IOneCcsFlashBoardProps, IOneCcsFlashBoardState, {}> {
  constructor(props: IOneCcsFlashBoardProps) {
    super(props);
    this.state = {
      ourFocus: "",
      description: [],
      carouselElements: [],
      vertical:"",
    };
    this._Horizontal=this._Horizontal.bind(this);
    this.vertical=this.vertical.bind(this);
  }

  public async componentDidMount() {
    this._ApplicationCategory();
    console.log(this.props.backGround);
    
  }

  private _ApplicationCategory = () => {
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.listName).items.get().then(results => {

      for (let i = 0; i < results.length; i++) {
        firstItem.push({
          Category: '<b>' + results[i].Category + '</b>' + ": " + results[i].Title,
          verCatogary:'<b>' + results[i].Category + '</b>' + ": <br/>" + results[i].Title,
        });
        console.log(firstItem);
      }
      //for horzontal
      for (let i = 0; i < firstItem.length - 1; i++) {
        let temp = firstItem[i].Category;
        firstItem[i].Category = firstItem[i + 1].Category; // firstItem[i].Category = temp;        
        final = temp + this.props.seperator + firstItem[i + 1].Category;      
        console.log(final);
      }
      //for vertical
      for (let i = 0; i < firstItem.length - 1; i++) {
        let temp = firstItem[i].verCatogary;
        firstItem[i].Category = firstItem[i + 1].verCatogary;    
      //for vertical
        verticalItem = temp + '<br/><br/>' + firstItem[i + 1].verCatogary + '<br/>' ;
        console.log(final);
      }
      this.setState({
        ourFocus: final,
        description: results,
        vertical:verticalItem,
      });
      console.log(results);
    });

  }

private _Horizontal(){  
  if(this.props.horizontal){
        return(
     
  <div style={{color:this.props.textColor}} dangerouslySetInnerHTML={{ __html: this.state.ourFocus }}></div>
    );
  }  
}
private vertical () {
  if(this.props.vertical){
    return(     
      //this.state.description.map((cat)=>{
        //return( 
        <div dangerouslySetInnerHTML={{ __html: this.state.vertical}} ></div>);       
      //})               
     //);     
  }
}


  public render(): React.ReactElement<IOneCcsFlashBoardProps> {

    return (
      <div >
         <div>
        <Marquee>
        {this._Horizontal()}
        </Marquee>
        </div>
        
        <div className={styles.marquee} style={{width:this.props.width}} >        
        {this.vertical()}     
        </div>
     
      </div>
    );
  }

}
