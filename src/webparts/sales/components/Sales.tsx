// import * as React from 'react';
// import styles from './Sales.module.scss';
// import type { ISalesProps } from './ISalesProps';
// import { escape } from '@microsoft/sp-lodash-subset';

// export default class Sales extends React.Component<ISalesProps> {
//   public render(): React.ReactElement<ISalesProps> {
//     const {
//       description,
//       isDarkTheme,
//       environmentMessage,
//       hasTeamsContext,
//       userDisplayName
//     } = this.props;

//     return (
//       <section className={`${styles.sales} ${hasTeamsContext ? styles.teams : ''}`}>
//         <div className={styles.welcome}>
//           <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
//           <h2>Well done, {escape(userDisplayName)}!</h2>
//           <div>{environmentMessage}</div>
//           <div>Web part property value: <strong>{escape(description)}</strong></div>
//         </div>
//         <div>
//           <h3>Welcome to SharePoint Framework!</h3>
//           <p>
//             The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
//           </p>
//           <h4>Learn more about SPFx development:</h4>
       
//         </div>
//       </section>
//     );
//   }
// }


import * as React from 'react';
import type { ISalesProps } from './ISalesProps';
import { HashRouter } from 'react-router-dom';
import RoutesItems from './Navigation/Routesitems';
import { SPHttpClient } from '@microsoft/sp-http';
import { hideLoader, showLoader } from './Shared/Loader';
import { ActionStatus } from './Constants/Contants';
import { showToast } from './Shared/Toaster';
import { ToastContainer } from 'react-toastify';
import { forEach } from 'lodash';
import NavBar from './Navigation/NavBar.components';

export default class Sales extends React.Component<ISalesProps> {

  public state = {
    isAuthorized: false,
    currentUserGroups: [],
    siteURL: '',
    webAbsoluteURL: '',
    currPlantTitle: '',
    isWCM: false,
    FormAccessConfiguration: {}
  }
  private siteURL = this.props.spContext.siteAbsoluteUrl;

  public componentDidMount() {
  
    this.removeExtraClasses();
  }



  private removeExtraClasses() {
    var workbenchElement = document.getElementById("workbenchPageContent");
    let wbClass = workbenchElement?.classList.value;
    workbenchElement?.classList.remove(wbClass ? wbClass : "");
    // this.removeAll();
    workbenchElement?.addEventListener("click", this.removeAll);

  }

  private removeAll = () => {
    var workbenchElement = document.getElementById("workbenchPageContent");
    workbenchElement?.removeEventListener("click", this.removeAll);

    var canvasComponent1 = document.getElementsByClassName("CanvasZoneContainer");
    forEach(canvasComponent1, element => {
      let eleClass = element.classList.value;
      let eleClassArr = eleClass.split(" ");

      eleClassArr.forEach((elem: string) => {
        element.classList.remove(elem.trim());
      });
    })
    var canvasComponent1 = document.getElementsByClassName("CanvasZone");
    forEach(canvasComponent1, element => {
      let eleClass = element.classList.value;
      let eleClassArr = eleClass.split(" ");

      eleClassArr.forEach((elem: string) => {
        element.classList.remove(elem.trim());
      });
    })
    var canvasComponent1 = document.getElementsByClassName("CanvasSection");
    forEach(canvasComponent1, element => {
      let eleClass = element.classList.value;
      let eleClassArr = eleClass.split(" ");

      eleClassArr.forEach((elem: string) => {
        element.classList.remove(elem.trim());
      });
    })
  }

  private onError = () => {
    showToast("error", ActionStatus.Error);
    hideLoader();
  }

  public render() {
    return (
      <React.Fragment>
        <ToastContainer />
        <HashRouter>
          <div className='menu-hide wrapper d-flex align-items-stretch' id="sideMenuNav">
             <NavBar {...this.props} {...this.state} />
             <RoutesItems {...this.props} {...this.state} />
          </div>
        </HashRouter>
      </React.Fragment>
    )
  }
}

