import * as React from 'react';
import styles from './OneCcsApplicationDashboard.module.scss';
import { IOneCcsApplicationDashboardProps } from './IOneCcsApplicationDashboardProps';
import './dashboard.css';
import { ICamlQuery, Item, sp } from "@pnp/sp/presets/all";
import { Modal } from '@fluentui/react';
import { Callout, Dialog, FontWeights, getTheme, IButtonStyles, Icon, IconButton, IIconProps, IPanelStyles, ISearchBoxStyles, ITextFieldStyles, ITooltipHostStyles, Link, mergeStyleSets, Panel, PanelType, PrimaryButton, SearchBox, TextField, TooltipHost } from 'office-ui-fabric-react';
import "@pnp/polyfill-ie11";
import 'polyfill-array-includes';
const calloutProps = { gapSpace: 0 };
const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
const modelProps = {
  isBlocking: true,
  topOffsetFixed: true,
};
const theme = getTheme();
const cancelIcon: IIconProps = { iconName: "cancel" };
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: " 50%",
    height: "73%",
    minWidth: "30%",
    overflowY: 'hidden'
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
    //overflowY: 'hidden',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
});
const innerContentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: " 50%",
    height: "29%",
    minWidth: "30%"

  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      //borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'left',
      fontWeight: FontWeights.semibold,
      padding: '12px 12px 14px 24px',
    },
  ],
  body: {
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
const iconButtonStyles: Partial<IButtonStyles> = {
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
const panelStyles: Partial<IPanelStyles> = {
  root: {
    width: 382,
    // marginTop: "10%",
    // marginRight: "10%",
    // height: "120%",
    inset: "93px 358px 19px 40%"
  },
};
const searchBoxStyles: Partial<ITextFieldStyles> = { root: { borderRadius: "19px", width: "94%", padding: '0 0 25px 0' } };
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
}
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
      catOpen: 'show',
      catName: "",
      divItems: [],
    };
    this._modalClose = this._modalClose.bind(this);
    this._panelClose = this._panelClose.bind(this);
    this.groupedCategory = this.groupedCategory.bind(this);
    this.bindInternalApplication = this.bindInternalApplication.bind(this);
    this._onOpenSearchPanel = this._onOpenSearchPanel.bind(this);
  }
  public async componentDidMount() {
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.listName).items.get().then(results => {
      this.setState({
        categoryChoices: results,
        categoryItems: results
      });

      console.log(this.state.categoryChoices);
      console.log(results);
    });

    this._searchItems();
  }
  private _searchItems = () => {
    sp.web.getList(this.props.siteUrl + "/Lists/InternalApplications").items.select("ApplicationCategory/ID,ApplicationCategory/Title,Title,ExternalLinkIcons,Link").expand("ApplicationCategory").get().then(search => {
      let grouping = search.reduce((r, a) => {
        r[a.ApplicationCategory.ID] = [...r[a.ApplicationCategory.ID] || [], a];
        return r;
      }, {});
      this.setState({
        searchtext: Object.values(grouping),
        items: Object.values(grouping),

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
    let filter = this.state.searchtext.filter(item => {
      return Object.keys(item).some(key => {
        return item[key]["Title"].toString().toLowerCase().trim().includes(text.length >= 2 ? text.toString().toLowerCase().trim() : "");
      });
    });

    this.setState({
      items: filter,
    });

  }
  public groupedCategory(i) {
    const myDiv = [];

    for (let j = 0; j < this.state.items[i].length; j++) {
      myDiv.push(
        <div className={styles.squareCat} >
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
    const emojiIcon: IIconProps = { iconName: 'Search' };
    return (
      <div className={styles.dasboard} style={{ width: "40%" }} >
        <div style={{ fontStyle: "bold", fontSize: 20, textAlign: 'left' }}>{this.props.description}</div>
        <IconButton iconProps={emojiIcon} ariaLabel="Emoji" onClick={this._onOpenSearchPanel} />Search

        < div className={styles.gridContainer} >
          {
            this.state.categoryItems.map((cat, key) => {
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
            })
          }

        </div >

        <div id="modal" >
          <Modal
            isOpen={this.state.shouldhide}
            onDismiss={() => this._modalClose}
            containerClassName={innerContentStyles.container}
          >
            <div className={contentStyles.header}>
              <div style={{ fontSize: "20px" }}>{this.state.applicationCategory}</div>
              <IconButton
                styles={iconButtonStyles}
                iconProps={cancelIcon}
                ariaLabel="Close popup modal"
                id="close"
                onClick={() => this.setState({ shouldhide: false })} />
            </div>
            <div className={styles.internalApplicationItems} style={{ display: "flex", justifyContent: "space-evenly", padding: "15px 28px 23px 44px" }}>
              {this.state.internalApplications.map((intAppItems, key) => {
                return (
                  <div >
                    <div onClick={() => this.loadLink(intAppItems)} >
                      <div style={{
                        background: "#f1e6e6", width: "48px", borderRadius: "20px", padding: "1px",
                        height: "60px"
                      }}>
                        <img style={{ width: 50, height: 59 }} src={intAppItems.ExternalLinkIcons != null ? intAppItems.ExternalLinkIcons.Url : ""} />
                      </div>
                      <div style={{ padding: "0px 0 0 30px", wordBreak: 'break-all', width: 90 }} className={styles.textbreak}>{intAppItems.Title}</div>
                    </div>
                  </div>
                );
              }
              )}
            </div>
          </Modal>
          <Modal
            isOpen={this.state.callOut}
            onDismiss={() => this._panelClose}
            containerClassName={contentStyles.container}
          >
            <div className={contentStyles.header}>
              <IconButton
                styles={iconButtonStyles}
                iconProps={cancelIcon}
                ariaLabel="Close popup modal"
                id="close"
                onClick={() => this.setState({ callOut: false })} />
            </div>
            <TextField
              onChange={this._onFilter}
              style={{ borderRadius: "14px", width: "84%", padding: "0px 0px 0 39px" }}
              autoComplete='off'
              placeholder="Find Applications."
            />
            <div style={{ padding: '10px 18px 21px 20px' }}>
              {
                this.state.items.map((cat, key) => {
                  return (
                    <div style={{ fontWeight: 600 }}>
                      <h4>{Object.values(cat)[0]["ApplicationCategory"]["Title"]}</h4>
                      <div style={{ display: 'flex' }} >
                        {this.groupedCategory(key)}
                      </div>
                    </div>
                  );
                })
              }

            </div>

          </Modal>
        </div>

      </div >

    );
  }
}


