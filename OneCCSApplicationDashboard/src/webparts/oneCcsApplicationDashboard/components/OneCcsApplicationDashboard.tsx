import * as React from 'react';
import styles from './OneCcsApplicationDashboard.module.scss';
import { IOneCcsApplicationDashboardProps } from './IOneCcsApplicationDashboardProps';
import './dashboard.css';
import { sp } from "@pnp/sp/presets/all";
import { Modal } from '@fluentui/react';
import { FontWeights, getTheme, IButtonStyles, Icon, IconButton, IIconProps, IPanelStyles, ISearchBoxStyles, ITextFieldStyles, ITooltipHostStyles, Link, mergeStyleSets, Panel, PanelType, PrimaryButton, SearchBox, TextField, TooltipHost } from 'office-ui-fabric-react';
import "@pnp/polyfill-ie11";
import 'polyfill-array-includes';
import * as _ from 'lodash';


const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: " 80%",
    height: "73%",
    minWidth: "30%",
    overflow: 'hidden'
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      //borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      padding: '12px 12px 14px 24px',
    },
  ],
  body: {
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    overflow: 'visible',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
});
const searchBoxStyles: Partial<ISearchBoxStyles> = { root: { borderRadius: "6rem", width: "18rem", height: "2rem" } };
export interface IOneCcsApplicationDashboardState {
  internalApplications: any[];
  shouldhide: boolean;
  categoryChoices: any[];
  categoryItems: any[];
  applicationCategory: string;
  icon: string;
  callOut: boolean;
  searchtext: any[];
  items: any[];
  transition: string;
  catOpen: string;
  catName: string;
  divItems: any[];
  searchDiv: string;
  backButton: string;

}
let sortedArray = [];
export default class OneCcsApplicationDashboard extends React.Component<IOneCcsApplicationDashboardProps, IOneCcsApplicationDashboardState, {}> {
  constructor(props: IOneCcsApplicationDashboardProps) {
    super(props);
    this.state = {
      internalApplications: [],
      shouldhide: false,
      categoryChoices: [],
      categoryItems: [],
      applicationCategory: "",
      icon: "",
      callOut: false,
      searchtext: [],
      items: [],
      transition: 'none',
      catOpen: "",
      catName: "",
      divItems: [],
      searchDiv: 'none',
      backButton: 'none',

    };
    this._modalClose = this._modalClose.bind(this);
    this._panelClose = this._panelClose.bind(this);
    this.groupedCategory = this.groupedCategory.bind(this);
    this.bindInternalApplication = this.bindInternalApplication.bind(this);
    this._onOpenSearchPanel = this._onOpenSearchPanel.bind(this);
    this._ApplicationCategory = this._ApplicationCategory.bind(this);
  }
  public async componentDidMount() {
    this._ApplicationCategory();
    this._searchItems();
  }
  private _ApplicationCategory = () => {
    this.setState({
      catOpen: "",
      transition: 'none',
      searchDiv: 'none',
      backButton: 'none',
      internalApplications: [],
    });
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.listName).items.get().then(results => {
      this.setState({
        categoryChoices: results,
        categoryItems: results,

      });

      console.log(this.state.categoryChoices);
      console.log(results);
    });

  }
  private _searchItems = () => {
    sp.web.getList(this.props.siteUrl + "/Lists/InternalApplications").items.select("ApplicationCategory/ID,ApplicationCategory/Title,Title,ExternalLinkIcons,Link").expand("ApplicationCategory").get().then(search => {
      // let grouping = search.reduce((r, a) => {
      //   r[a.ApplicationCategory.ID] = [...r[a.ApplicationCategory.ID] || [], a];
      //   return r;
      // }, {});
      sortedArray = _.orderBy(search, 'Title', ['asc']);
      this.setState({
        searchtext: sortedArray,
        items: sortedArray,

      });
    });
    console.log(this.state.searchtext);
  }
  private bindInternalApplication = (cat, key) => {
    this.setState({
      shouldhide: true,
      applicationCategory: cat.Title,
      icon: cat.IconName,
      transition: "",
      catOpen: 'none',
      backButton: "",
    });
    sp.web.getList(this.props.siteUrl + "/Lists/InternalApplications").items.filter("ApplicationCategoryId eq '" + cat.ID + "'").get().then(iAppItems => {
      this.setState({
        internalApplications: iAppItems,
      });
      console.log(this.state.internalApplications);
      console.log(iAppItems[0].Link.Url);
    });
  }
  private loadLink = (intAppItems) => {

    return (
      window.open(intAppItems.Link.Url)
    );
  }
  private _modalClose = () => {
    this.setState({
      shouldhide: false,
      callOut: false,
      internalApplications: [],
    });
  }
  private _onOpenSearchPanel() {
    this._searchItems();
    this.setState({
      callOut: true
    });
  }
  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    if (text == "") {
      this._ApplicationCategory();
    }
    else {
      this.setState({
        items: text ? this.state.searchtext.filter(i => i.Title.toLowerCase().indexOf(text.toString().toLowerCase()) > -1) : this.state.searchtext,
        searchDiv: "",
        catOpen: 'none',
        transition: 'none',

      });
    }
  }
  public groupedCategory(i) {
    const myDiv = [];

    for (let j = 0; j < this.state.items[i].length; j++) {
      myDiv.push(
        <div>
          <img style={{ width: 50, height: 59, padding: "0% 0% 0% 12%" }} src={this.state.items[i][j].ExternalLinkIcons != null ? this.state.items[i][j].ExternalLinkIcons.Url : ""} onClick={() => window.open(this.state.items[i][j].Link.Url)} />
          <div className={styles.textbreak} style={{ padding: "0px 0px 0px 15px", wordBreak: 'break-all', width: 90, textAlign: "match-parent" }}>{this.state.items[i][j].Title}</div>
        </div>
      );

    }

    return myDiv;
  }
  private _panelClose = () => {
    this.setState({
      callOut: false,
    });

  }

  public render(): React.ReactElement<IOneCcsApplicationDashboardProps> {
    const CatIcon = () => <Icon iconName={this.state.icon} />;
    const backIcon: IIconProps = { iconName: 'Back' };
    return (
      <>
        <div className={styles.dasboard} style={{ width: "20rem" }}>
          <div style={{ fontStyle: "bold", fontSize: 20, textAlign: 'left' }}>{this.props.description}</div>
          <SearchBox placeholder="Type application name" styles={searchBoxStyles} onSearch={newValue => console.log('value is ' + newValue)} onChange={this._onFilter} />
          <div style={{ display: this.state.backButton }}> <IconButton iconProps={backIcon} ariaLabel="Emoji" onClick={() => this._ApplicationCategory()} /></div>
          <div style={{ display: this.state.catOpen }}>
            <div className={styles.gridContainer}>
              {this.state.categoryItems.map((cat, key) => {
                return (
                  <div>
                    <div className={styles.squareCat} style={{ background: cat.BackgroundColor }} onClick={() => this.bindInternalApplication(cat, key)}>
                      <div>
                        <i className={styles['msIcon']} aria-hidden="true"><Icon iconName={cat.IconName} /></i>
                      </div>
                    </div>
                    <div style={{ padding: '10px' }}> {cat.Title}</div>
                  </div>

                );
              })}
            </div>
          </div>
          {/* Inner application binding */}
          <div style={{ display: this.state.transition }}>
            <div className={contentStyles.header}>
              <div style={{ fontSize: "20px" }}>{this.state.applicationCategory}</div>
            </div>
            <div className={styles.gridContainer1}>
              {this.state.internalApplications.map((intAppItems, key) => {
                return (
                  <div style={{ padding: "0rem 1rem 0rem 0rem" }}>
                    <div onClick={() => this.loadLink(intAppItems)} >
                      <div className={styles.squareCat1}>

                        <img src={intAppItems.ExternalLinkIcons != null ? intAppItems.ExternalLinkIcons.Url : ""} />
                      </div>
                      <div style={{ padding: '10px' }}> {intAppItems.Title}</div>
                    </div>
                  </div>
                );
              }
              )}
            </div>
          </div>
          {/* Search items */}
          <div style={{ display: this.state.searchDiv }} >

            {this.state.items.map((searchItems, key) => {
              return (
                <div style={{ padding: "1rem 0 0 0" }}>
                  <table style={{ cursor: "pointer" }}>
                    <tr onClick={() => this.loadLink(searchItems)}  >
                      <td><img src={searchItems.ExternalLinkIcons != null ? searchItems.ExternalLinkIcons.Url : ""} /></td>
                      <td style={{ padding: "0 0 0 1rem" }} ><label style={{ fontWeight: "bold" }}>{searchItems.Title}</label>
                        <br></br>
                        <label style={{ cursor: "pointer" }}>Category : {searchItems.ApplicationCategory.Title} </label></td>
                    </tr>
                  </table>
                </div>
              );
            })}

          </div>
        </div></>

    );
  }

}


