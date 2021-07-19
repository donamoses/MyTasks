import * as React from 'react';
import styles from './OneCcsApplicationDashboard.module.scss';
import { IOneCcsApplicationDashboardProps } from './IOneCcsApplicationDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import './dashboard.css';
import { Item, sp } from "@pnp/sp/presets/all";
//import { IconButton, Modal } from 'office-ui-fabric-react';
import { Modal } from '@fluentui/react';
import { FontWeights, getTheme, IButtonStyles, Icon, IconButton, IIconProps, ITextFieldStyles, Link, mergeStyleSets, TextField } from 'office-ui-fabric-react';
import { Container, GridList, GridListTile, GridListTileBar } from '@material-ui/core';
const theme = getTheme();
const cancelIcon: IIconProps = { iconName: "cancel" };
{/* <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet"></link> */ }
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: " 30 %",
    height: "73 %",

  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      borderTop: `4px solid ${theme.palette.themePrimary}`,
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
const searchStyles: Partial<ITextFieldStyles> = {
  root: {
    width: "88%",
    BorderBottom: "groove",
  },

};
export interface IOneCcsApplicationDashboardState {
  internalApplications: any[];
  shouldhide: boolean;
  categoryChoices: any[];
  categoryItems: any[];
  applicationCategory: string;
  icon: string;

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
    };
    this._modalClose = this._modalClose.bind(this);
    this.bindInternalApplication = this.bindInternalApplication.bind(this);
  }

  public async componentDidMount() {

    // sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.listName).fields.getByInternalNameOrTitle("Category").select("Choices").get().then(results => {
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.listName).items.get().then(results => {
      this.setState({
        categoryChoices: results,
        categoryItems: results
      });

      console.log(this.state.categoryChoices);
      console.log(results);
    });
  }
  private _modalClose = () => {

    this.setState({ shouldhide: false });
  }
  private bindInternalApplication = (cat, key) => {

    this.setState({
      shouldhide: true,
      applicationCategory: cat.Title,
      icon: cat.IconName,
    });

    sp.web.getList(this.props.siteUrl + "/Lists/InternalApplications").items.filter("ApplicationCategoryId eq '" + cat.ID + "'").get().then(iAppItems => {
      this.setState({
        internalApplications: iAppItems,

      });
      console.log(this.state.internalApplications);
      console.log(iAppItems[0].Link.Url);
    });
  }
  private loadLink = (intAppItems: { Link: { Url: string; }; }, key: number) => {
    return (
      // window.location.replace(intAppItems.Link.Url)
      window.open(intAppItems.Link.Url)

    );

  }
  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {

    this.setState({
      categoryItems: text ? this.state.categoryChoices.filter(i => i.name.toLowerCase().indexOf(text) > -1) : this.state.categoryChoices,
    });
  }
  public render(): React.ReactElement<IOneCcsApplicationDashboardProps> {
    const CatIcon = () => <Icon iconName={this.state.icon} />;
    return (
      <div className={styles.dasboard} style={{ width: "50%" }}>
        <TextField
          // className={exampleChildClass}
          label="Search"
          onChange={this._onFilter}
          styles={searchStyles}
        />
        <div className={styles.gridContainer}>
          {
            this.state.categoryItems.map((cat, key) => {
              return (
                <div>
                  <div className={styles.squareCat} style={{ background: cat.BackgroundColor }} onClick={() => this.bindInternalApplication(cat, key)}>
                    <div>
                      {/* <IconButton
                        iconProps={{ iconName: cat.IconName }} /> */}
                      <i className={styles['ms-Icon']} aria-hidden="true"><Icon iconName={cat.IconName} /></i>
                      {/* <Icon iconName={cat.IconName} /> */}
                    </div>
                  </div>
                  <div style={{ padding: '10px' }}> {cat.Title}</div>
                </div>

              );
            })
          }
        </div>

        <div id="modal" >
          <Modal
            isOpen={this.state.shouldhide}
            onDismiss={() => this._modalClose}
            containerClassName={contentStyles.container}
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
            <div className={styles.internalApplicationItems} style={{ display: "flex", justifyContent: "space-evenly" }}>
              {this.state.internalApplications.map((intAppItems, key) => {
                return (
                  <div >
                    <div onClick={() => this.loadLink(intAppItems, key)} >
                      <div style={{
                        background: "#f1e6e6", width: "48px", borderRadius: "20px", padding: "20px",
                        height: "60px"
                      }}>
                        <img style={{ width: 50, height: 59 }} src={intAppItems.ExternalLinkIcons != null ? intAppItems.ExternalLinkIcons.Url : ""} />
                      </div>
                      <h3 style={{ padding: "0px 0 0 25px" }}>{intAppItems.Title}</h3>
                    </div>
                  </div>
                );
              }
              )}
            </div>

          </Modal>
        </div>



      </div >
    );
  }
}