import * as React from 'react';
import styles from './ModernWorkSpace.module.scss';
import { IModernWorkSpaceProps } from './IModernWorkSpaceProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { IModernWorkSpaceState } from './IModernWorkSpaceState';
import { ICheckboxInput } from './model/ICheckboxInput';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
// import ProgressBar from 'react-customizable-progressbar';
import { SearchBox } from 'office-ui-fabric-react/lib';
import { HttpClientResponse, AadHttpClient } from '@microsoft/sp-http';
require('../../../../node_modules/jquery/dist/jquery.min.js');
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
import { PISearchResult, PSearchResult } from './model/ISearchResult';
import { PISearchResults } from './model/SPResponse';
import InlineSearchResults from './InlineSearchResults/InlineSearchResults';
import ListSearchProductCard from './InlineSearchResults/ListSearchProductCard';
import SPSearchService from './model/SPSearchService';
import { LineGraphInData, Activity, GroupedActivity, AggregatedActivity, LineGraphData, PieGraphData, PieData,CollaborationActivity} from "./interfaces/Objects";

import { Carousel, CarouselButtonsDisplay, CarouselButtonsLocation } from "@pnp/spfx-controls-react/lib/Carousel";
// import { IoMdAttach } from 'react-icons/io';
import { SPEvents, ISPEventItem} from './interfaces/ISPEventBirthday';
import * as  moment from 'moment';
import {FontAwesomeIcon}  from '@fortawesome/react-fontawesome';
import { config } from '@fortawesome/fontawesome-svg-core';
import { parse, toSeconds } from 'iso8601-duration';
config.autoA11y = true;
import { TextField, Slider, PrimaryButton, IPersonaProps, IPersonaSharedProps, Persona, PersonaSize, Callout, DirectionalHint, HighContrastSelector, ImageFit, initializeIcons, styled } from 'office-ui-fabric-react/lib';
import { escape } from '@microsoft/sp-lodash-subset';
import { Environment, EnvironmentType, ServiceScope } from '@microsoft/sp-core-library';
import {ISPSearchResult, LobLink, ILobLinkItems} from './interfaces/ISPSearchResult';
import { ISearchResults, ICells, ICellValue, ISearchResponse } from './interfaces/ISearchService';
import {
  MSGraphClient,
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { ChartControl, ChartType } from '@pnp/spfx-controls-react/lib/ChartControl';
// import { Carousel } from 'react-responsive-carousel';
// import { Carousel } from 'react-responsive-carousel';
// import "react-responsive-carousel/lib/styles/carousel.min.css";
import { Rerousel } from 'rerousel';

import { IGraphMail, IGraphMailItems, GraphMail } from './interfaces/IGraphMail';
import { IGraphMyTeam, IGraphMyTeamItems, GraphMyTeam, IGraphTeamMessage, IGraphTeamMessageItems, GraphTeamMessage } from './interfaces/IGraphMyTeam';
import { IGraphUserProfile, IGraphUserProfileItems, GraphUserProfile, ISPUsers, SPUsers } from './interfaces/IGraphUserProfile';
import { IGraphDriveFile, IGraphDriveFileItems, GraphDriveFile } from './interfaces/IGraphDriveFile';
import { ISPEvent, ISPEventItems, SPEvent, myOutlookEvent } from './interfaces/ISPEvent';
import { CarouselImageService } from './model/CarouselImageService';
import { IDataService } from './model/IDataService';
import { ISPAnnouncement, ISPAnnouncementItems, SPAnnouncement } from './interfaces/ISPAnnouncement';
import { ILeaveRequestItems, LeaveRequest, ILoanRequestItems, LoanRequest, IPettyCashItems, PettyCash, ISalaryAdvItems, SalaryAdv, completedProcess, ICompletedProcess, aggProcess } from './interfaces/IProcessStage';
import { IEngageProgressItem, EngageProgress } from './interfaces/IEngageProgress';
import { ITimeSheetItem,  TimeSheet } from './interfaces/ITimeSheet';
import { ICarouselImage,ICarouselImageFile,ICarouselImageRLink } from './interfaces/IGraphDriveFile'; 
import { faWindowRestore } from '@fortawesome/free-solid-svg-icons';
import { useRef } from 'react';

import * as Data from './sampleData';
import { ResponsiveBar } from '@nivo/bar';
import { ResponsivePie } from '@nivo/pie';
import * as reactIframe from 'react-iframe';
import { DynamicsTask } from './interfaces/IDynamicsTask';

function formatDateTime(dateTime) { 
  return moment.utc(dateTime).local().format('DD-MMMM-YY');
}

function formatDate(dateTime) { 
  return moment.utc(dateTime).local().format('DD-MMMM');
}

function App() {
  const ref = useRef(null);

}

const options: ICheckboxInput[] = [
  { ID: 1, Title: 'Microsoft Teams' },
  { ID: 2, Title: 'Human360' },
  { ID: 3, Title: 'Dynamics 365' },
  { ID: 4, Title: 'Plant Maintenance' },
  { ID: 5, Title: 'Virtual Receptionist' },
  { ID: 6, Title: 'Windows Virtual Desktop' },
];

// Used to add spacing between checkboxes
const stackTokens = { childrenGap: 10 };

const options1: ICheckboxInput[] = [
  { ID: 1, Title: 'NOSAK Group' },
  { ID: 2, Title: 'Aiico Insurance' },
  { ID: 3, Title: 'Eunisell' },
  { ID: 4, Title: 'MTN' },
  { ID: 5, Title: 'Central Bank of Nigeria' },
  { ID: 6, Title: 'Africa Alliance Insurance Plc' },
];

// Used to add spacing between checkboxes
const stackTokens1 = { childrenGap: 10 };

const options2: ICheckboxInput[] = [
  { ID: 1, Title: 'Oil and Gas' },
  { ID: 2, Title: 'Telecommunication' },
  { ID: 3, Title: 'Manufacturing' },
  { ID: 4, Title: 'HealthCare' },
  { ID: 5, Title: 'Agriculture' },
  { ID: 6, Title: 'Financial Service' },
];

// Used to add spacing between checkboxes
const stackTokens2 = { childrenGap: 10 };

const options3: ICheckboxInput[] = [
  { ID: 1, Title: 'Energy and Power' },
  { ID: 2, Title: 'Education' },
  { ID: 3, Title: 'Logistics' },
  { ID: 4, Title: 'Tourism' },
  { ID: 5, Title: 'Commercial/Retail' },
  { ID: 6, Title: 'Media and Entertainment' },
];

// Used to add spacing between checkboxes
const stackTokens3 = { childrenGap: 10 };

export default class ModernWorkSpace extends React.Component<IModernWorkSpaceProps, IModernWorkSpaceState> {
  private dataCenterServiceInstance: IDataService;
  //private carouselInterval;
  constructor(props:  IModernWorkSpaceProps, state: IModernWorkSpaceState){
    super(props);
    this.state = {
      isConnected: false,
      isOnTeams: false,
      notificationCount: 0,
      allUser: [],
      myRecentUsers: [],
      myProfile: null,
      mailMessageArr: [],
      mailMessageCount: 0,
      myTeamGroupsArr: [],
      selectedTeam: "myTeam_1",
      selectedTeamMessages: [],
      myRecentDriveFiles: [],
      docsSharedWithMe: [],
      companyEvents: [],
      staffBirthdays: [],
      staffAnniversary: [],
      myCalendar: [],
      mySPSites: [],
      spUsers: [],
      spAnnouncements: [],
      myPlannerTasks: [],
      spFAQ: [],
      lineExtn: [],
      searchExtn: "",
      employeeCount: [],
      searchstatus: false,
      LOBLinkList: [], 
      searchText: "", 
      item: [],
      isCalloutVisible: false,
      myLeaveRequest: [],
      loanRequests: [],
      pettyCashApp: [],
      salaryAdvApp: [],
      allProcessStages: [],
      completedProcesses: [],
      aggregratedProcesses: [],
      searchProcessStage: "",
      allEngagements: [],
      searchEngage: "",
      allTimeSheet: [],
      searchTimeSheet: "",
      PsearchText: "",
      PisCalloutVisible: false,
      PsearchResult: [],
      resultToDisplay: [],
      Psearchstatus: false,
      value: 'See More',
      productChecked: [],
      customerChecked: [],
      sectorChecked: [],
      industryChecked: [],
      GroupActivityList: {Email: [], Focus: [], Meeting: [],Call: [], Chat: []},
      AggregatedActivityList: [],
      CollaborationActivityList: {Email: [], Focus: [], Meeting: [], "Chat/Call": []},
      barChartData: [],
      pieChartData: [],
      imageURLs: [],
      resourceLink: [],
      carouselItems: [],
      carouselItemElements: [],
      currentCarouselItemElement: <div></div>,
      currentCarouselItemIndex: 0,
      carouselCanMoveNext: false,
      carouselCanMovePrev: false
    };

    let serviceScope: ServiceScope = this.props.serviceScope;
  
    // this.dataCenterServiceInstance = serviceScope.consume(CarouselImageService.serviceKey);
    // this.dataCenterServiceInstance.getImages('Gallery Slide').then((carouselItems: ICarouselImage[]) => {
    //   console.log(carouselItems);
    //   this.setState({
    //     carouselItems: carouselItems,
    //     carouselItemElements: carouselItems.map(imageList => (
    //     <div style={{display:"flex", height:"100%"}}>
    //       <a href={imageList.ResourceLink} target='_blank'>
    //         <img style={{width:'100%', height:'100%', objectFit:"contain"}} src={imageList.FileRef} />
    //       </a> 
    //     </div>)
    //     )
    //    // imageURLs: this._getGalleryImages(carouselItems) ,
    //     //resourceLink: this._getGalleryResourseLink(carouselItems)

    //   });
    // });

  }
  public _msGraphFactory: Promise<MSGraphClient> = this.props.context.msGraphClientFactory.getClient();
  
  private _menuButtonElement = React.createRef<HTMLDivElement>();

  
  public render(): React.ReactElement<IModernWorkSpaceProps> {
    if(Environment.type == EnvironmentType.ClassicSharePoint || Environment.type == EnvironmentType.SharePoint){

      initializeIcons();
      const analyticsBarChart1 = <ResponsiveBar
        data={this.state.barChartData}
        keys={[ 'Collab', 'Focus' ]}
        indexBy="DAY"
        margin={{ top: 40, right: 20, bottom: 40, left: 50 }}
        padding={0.3}
        enableGridY={true}
        gridYValues={5}
        // colors={{ scheme: 'nivo' }}
        defs={[
          {
            id: 'green',
            type: 'patternDots',
            background: '#64c8d8',
            color: '#64c8d8',
            size: 1,
            padding: 0,
            stagger: true
          },
          {
            id: 'blue',
            type: 'patternLines',
            background: '#5b94a5',
            color: '#5b94a5',
            rotation: 45,
            lineWidth: 6,
            spacing: 1
          }
        ]}
        fill={[
          {
            match: {
              id: 'Collab'
            },
            id: 'blue'
          },
          {
            match: {
              id: 'Focus'
            },
            id: 'green'
          }
        ]}
        borderColor={{ from: 'color', modifiers: [ [ 'darker', 1.6 ] ] }}
        axisTop={null}
        axisRight={null}
        axisBottom={{
          tickSize: 5,
          tickPadding: 5,
          tickRotation: 0,
          legend: <strong>DAY</strong>,
          legendPosition: 'middle',
          legendOffset: 15
        }}
        axisLeft={{
          tickSize: 5,
          tickPadding: 5,
          tickRotation: 0,
          /* tickValues: [0, 20, 40, 60],
          legend: 'Percent Usage',
          legendPosition: 'middle',
          legendOffset: -40 */
        }}
        enableLabel={false}
        // labelFormat={}
        legends={[
          {
            dataFrom: 'keys',
            anchor: 'top',
            direction: 'row',
            justify: false,
            // translateX: 0,
            translateY: -30,
            itemsSpacing: 1,
            itemWidth: 80,
            itemHeight: 20,
            itemDirection: 'left-to-right',
            itemOpacity: 0.85,
            symbolSize: 15,
            effects: [
              {
                on: 'hover',
                style: {
                  itemOpacity: 1
                }
              }
            ]
          }
        ]}
        animate={true}
        motionStiffness={90}
        motionDamping={15}
      />;

      const analyticsBarChart2 = <ResponsiveBar
        data={Data.barChartData2}
        keys={[ 'Meeting', 'Email' ]}
        indexBy="Department"
        margin={{ top: 80, right: 20, bottom: 60, left: 60 }}
        padding={0.3}
        enableGridY={true}
        gridYValues={5}
        // colors={{ scheme: 'nivo' }}
        defs={[
          {
            id: 'pink',
            type: 'patternDots',
            background: '#eb623e',
            color: '#eb623e',
            size: 1,
            padding: 0,
            stagger: true
          },
          {
            id: 'grey',
            type: 'patternLines',
            background: '#c7c9c8',
            color: '#c7c9c8',
            rotation: 45,
            lineWidth: 6,
            spacing: 1
          }
        ]}
        fill={[
          {
            match: {
              id: 'Meeting'
            },
            id: 'pink'
          },
          {
            match: {
              id: 'Email'
            },
            id: 'grey'
          }
        ]}
        borderColor={{ from: 'color', modifiers: [ [ 'darker', 1.6 ] ] }}
        axisTop={null}
        axisRight={null}
        axisBottom={{
          tickSize: 2,
          tickPadding: 5,
          tickRotation: 45,
          legend: <strong>Department</strong>,
          legendPosition: 'middle',
          legendOffset: 60
        }}
        axisLeft={{
          tickSize: 5,
          tickPadding: 5,
          tickRotation: 0,
          tickValues: [5, 10, 15, 20],
          legend: 'Hours Spent',
          legendPosition: 'middle',
          legendOffset: -40
        }}
        enableLabel={false}
        // labelFormat={}
        legends={[
          {
            dataFrom: 'keys',
            anchor: 'top',
            direction: 'row',
            justify: false,
            // translateX: 0,
            translateY: -60,
            itemsSpacing: 2,
            itemWidth: 80,
            itemHeight: 40,
            itemDirection: 'left-to-right',
            itemOpacity: 0.85,
            symbolSize: 20,
            effects: [
              {
                on: 'hover',
                style: {
                  itemOpacity: 1
                }
              }
            ]
          }
        ]}
        animate={true}
        motionStiffness={90}
        motionDamping={15}
      />;
    
      const analyticsBarChart3 = <ResponsiveBar
        data={this.state.aggregratedProcesses}
        layout="horizontal"
        // keys={[ 'Collab', 'Focus' ]}
        keys={[ 'Avg. Duration' ]}
        indexBy="Process"
        margin={{ top: 40, right: 20, bottom: 40, left: 75 }}
        padding={0.3}
        enableGridY={false}
        enableGridX={true}
        gridXValues={5}
        // colors={{ scheme: 'nivo' }}
        defs={[
          {
            id: 'green',
            type: 'patternDots',
            background: '#64c8d8',
            color: '#64c8d8',
            size: 1,
            padding: 0,
            stagger: true
          },
          {
            id: 'blue',
            type: 'patternLines',
            background: '#5b94a5',
            color: '#5b94a5',
            rotation: 45,
            lineWidth: 6,
            spacing: 1
          }
        ]}
        fill={[
          {
            match: {
              id: 'AvgDuration'
            },
            id: 'blue'
          }
        ]}
        borderColor={{ from: 'color', modifiers: [ [ 'darker', 1.6 ] ] }}
        axisTop={null}
        axisRight={null}
        axisBottom={{
          tickSize: 5,
          tickPadding: 5,
          tickRotation: 0,
          legend: <strong>DAY</strong>,
          legendPosition: 'middle',
          legendOffset: 15
        }}
        axisLeft={{
          tickSize: 5,
          tickPadding: 5,
          tickRotation: 0,
          /* tickValues: [0, 20, 40, 60],
          legend: 'Percent Usage',
          legendPosition: 'middle',
          legendOffset: -40 */
        }}
        enableLabel={false}
        // labelFormat={}
        legends={[
          {
            dataFrom: 'keys',
            anchor: 'top',
            direction: 'row',
            justify: false,
            // translateX: 0,
            translateY: -30,
            itemsSpacing: 1,
            itemWidth: 80,
            itemHeight: 20,
            itemDirection: 'left-to-right',
            itemOpacity: 0.85,
            symbolSize: 15,
            effects: [
              {
                on: 'hover',
                style: {
                  itemOpacity: 1
                }
              }
            ]
          }
        ]}
        animate={true}
        motionStiffness={90}
        motionDamping={15}
      />;
      
      const analyticsPieChart = (
        <ResponsivePie
          data={this.state.pieChartData}
          margin={{ top: 40, right: 10, bottom: 80, left: 10 }}
          colors={{ scheme: 'nivo' }}
          borderWidth={1}
          borderColor="#f7f7f7"
          enableRadialLabels={false}
          radialLabel={(e) => e.id+" ("+e.value+")"}
          radialLabelsSkipAngle={10}
          radialLabelsTextXOffset={6}
          radialLabelsTextColor="#333333"
          radialLabelsLinkOffset={0}
          radialLabelsLinkDiagonalLength={16}
          radialLabelsLinkHorizontalLength={24}
          radialLabelsLinkStrokeWidth={1}
          radialLabelsLinkColor={{ from: 'color' }}
          enableSlicesLabels={false}
          sliceLabel="id"
          slicesLabelsSkipAngle={10}
          slicesLabelsTextColor="#333333"
          animate={true}
          motionStiffness={90}
          motionDamping={15}
          defs={[
            {
              id: 'dots',
              type: 'patternDots',
              background: 'inherit',
              color: 'rgba(255, 255, 255, 0.3)',
              size: 4,
              padding: 1,
              stagger: true
            },
            {
              id: 'lines',
              type: 'patternLines',
              background: 'inherit',
              color: 'rgba(255, 255, 255, 0.3)',
              rotation: -45,
              lineWidth: 6,
              spacing: 10
            }
          ]}
          fill={[
            {
              match: {
                id: 'ruby'
              },
              id: 'dots'
            },
            {
              match: {
                id: 'c'
              },
              id: 'dots'
            },
            {
              match: {
                id: 'go'
              },
              id: 'dots'
            },
            {
              match: {
                id: 'python'
              },
              id: 'dots'
            },
            {
              match: {
                id: 'scala'
              },
              id: 'lines'
            },
            {
              match: {
                id: 'lisp'
              },
              id: 'lines'
            },
            {
              match: {
                id: 'elixir'
              },
              id: 'lines'
            },
            {
              match: {
                id: 'javascript'
              },
              id: 'lines'
            }
          ]}
          legends={[
            {
              anchor: 'bottom',
              direction: 'row',
              translateY: 56,
              itemWidth: 80,
              itemHeight: 18,
              itemTextColor: '#999',
              symbolSize: 18,
              symbolShape: 'circle',
              effects: [
                {
                  on: 'hover',
                  style: {
                    itemTextColor: '#000'
                  }
                }
              ]
            }
          ]}
        />
      );
      const analyticsPieChartSM = (
        <ResponsivePie
          data={this.state.pieChartData}
          margin={{ top: 10, right: 20, bottom: 10, left: 80 }}
          colors={{ scheme: 'nivo' }}
          borderWidth={1}
          borderColor="#f7f7f7"
          enableRadialLabels={false}
          radialLabel={(e) => e.id+" ("+e.value+")"}
          radialLabelsSkipAngle={10}
          radialLabelsTextXOffset={6}
          radialLabelsTextColor="#333333"
          radialLabelsLinkOffset={0}
          radialLabelsLinkDiagonalLength={16}
          radialLabelsLinkHorizontalLength={24}
          radialLabelsLinkStrokeWidth={1}
          radialLabelsLinkColor={{ from: 'color' }}
          enableSlicesLabels={false}
          sliceLabel="id"
          slicesLabelsSkipAngle={10}
          slicesLabelsTextColor="#333333"
          animate={true}
          motionStiffness={90}
          motionDamping={15}
          defs={[
            {
              id: 'dots',
              type: 'patternDots',
              background: 'inherit',
              color: 'rgba(255, 255, 255, 0.3)',
              size: 4,
              padding: 1,
              stagger: true
            },
            {
              id: 'lines',
              type: 'patternLines',
              background: 'inherit',
              color: 'rgba(255, 255, 255, 0.3)',
              rotation: -45,
              lineWidth: 6,
              spacing: 10
            }
          ]}
          fill={[
            {
              match: {
                id: 'ruby'
              },
              id: 'dots'
            },
            {
              match: {
                id: 'c'
              },
              id: 'dots'
            },
            {
              match: {
                id: 'go'
              },
              id: 'dots'
            },
            {
              match: {
                id: 'python'
              },
              id: 'dots'
            },
            {
              match: {
                id: 'scala'
              },
              id: 'lines'
            },
            {
              match: {
                id: 'lisp'
              },
              id: 'lines'
            },
            {
              match: {
                id: 'elixir'
              },
              id: 'lines'
            },
            {
              match: {
                id: 'javascript'
              },
              id: 'lines'
            }
          ]}
          legends={[
            {
              anchor: 'left',
              direction: 'column',
              translateX: -70,
              itemWidth: 60,
              itemHeight: 20,
              itemTextColor: '#999',
              symbolSize: 18,
              symbolShape: 'circle',
              effects: [
                {
                  on: 'hover',
                  style: {
                    itemTextColor: '#000'
                  }
                }
              ]
            }
          ]}
        />
      );

      const myRecentUser: JSX.Element[] = this.state.myRecentUsers.map((user, i) => {
        let userLink = `https://delve.office.com/?u=${user.id}&v=work`, linkToMail = `mailto:${user.userPrincipalName}?body=Hello%20${user.displayName}%0D%0A%0D%0A`;
        // if(i<12) 
        return (
        <a className={`ms-Grid-col ms-sm2 ms-md2 ms-lg2 ms-xl2 ms-xxl2 ms-xxxl2 ${styles.msGridCol} ${styles.msSm2}} ${styles.msMd2} ${styles.msLg2} ${styles.msXl2} ${styles.msXxl2} ${styles.msXxxl2} ${styles.col1memeber}`} target="_blank" href={userLink} title={user.displayName}>
          <div className={`${styles.col1memberImg}`} style={{backgroundImage: `url('${this.props.siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${user.userPrincipalName}')`}}></div>
        </a>
        );
      });

      const myMailElArr: JSX.Element[] = this.state.mailMessageArr.map(mailItem => {
        const mailTime: Date = mailItem.receivedDateTime;
        const timeWhen1 = this.utilityMethod.timeSince(mailTime);
        const timeWhen = moment(mailTime).fromNow();

        const timeFrom: string = `${this.utilityMethod.daysOfTheWeek[mailTime.getDay()]} ${mailTime.getDate()<8 ? "0" +mailTime.getDate() : mailTime.getDate()}/${mailTime.getMonth()<8 ? "0" +(mailTime.getMonth() + 1) : mailTime.getMonth() +1}`;
        
        return(
          <div className={`${styles.col2MailSection}`}>
            {/* <p className={styles.col2MailTime} title={mailItem.subject}>{timeWhen}</p> */}
            <input type="checkbox" id={mailItem.id}></input>
            <label className={styles.col2MailTime} title={mailItem.subject} htmlFor={mailItem.id}>{timeWhen}</label>
            <a className={`${styles.col2MailContainer}`} href={mailItem.webLink} target="_blank" onClickCapture={()=>this._outlookMailListener(mailItem.id)}>
                <div className={`${styles.col2MailIcon}`}>
                  <i className={`ms-Icon ms-Icon--Mail`}></i>
                </div>
                <div className={`${styles.col2MailDetails}`}>
                    <div>
                      <span>{`From: ${mailItem.from}`}</span>
                      <span>{timeFrom}</span>
                    </div>
                    <div className={`${styles.col2MailContent}`}>{mailItem.bodyPreview}</div>
                </div>
            </a>
          </div>
        );
      });

      const myTeamGroups: JSX.Element[] = this.state.myTeamGroupsArr.map((team, i) => (
        <div>
          <input type="radio" name="mySelectedTeam" id={`myTeam_${i+1}`} data-index={i} data-teamId={team.teamId} data-groupId={team.groupId} checked={this.state.selectedTeam === `myTeam_${i+1}`} onChange={(e) => this._teamSelectedListener(e)}></input>
          <label htmlFor={`myTeam_${i+1}`} title={team.displayName}>{team.displayName}</label>
        </div>
      ));
      const myTeamMessages: JSX.Element[] = this.state.selectedTeamMessages.map(message => {
        let fromUser = null;
        if(message.from && message.from.user && message.from.user.id){
          fromUser = this.state.allUser.filter(user => {
            if(!message.from){return false;}
            return user.id == message.from.user.id;
          })[0];
        }
        
        return(
        <div>
          <div className={styles.colTeamsConvoBy}>
              {fromUser ? 
                <img src={`${this.props.siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${fromUser.mail}`} alt={fromUser.displayName} title={fromUser.displayName}></img> :
                <img src={require(`./images/account.svg`)} alt={``} title={``}></img>
              }
          </div>
          <div className={styles.colTeamsConvoMessage}>
              {fromUser ? <p>{fromUser.displayName}</p> : <p>User</p>}
              <div dangerouslySetInnerHTML= {{__html: message.body.content}}></div>
              {/* ReactHtmlParser ( message.body.content ) */}
          </div>
        </div>
      );
      });

      const myRecentDoc: JSX.Element[] = this.state.myRecentDriveFiles.map(doc => {
        const docExtn = doc.name.split(".").pop().toLowerCase(), extns = ["csv", "docx", "dotx", "onetoc", "pdf", "potx", "pptx", "pub", "xlsx", "xltx", "zip", "png", "jpg"];
        
        let docClass = extns.indexOf(docExtn) !== -1 ? styles[`${docExtn}Doc0`] : styles.fileDoc0;
                
        return(
          <p className={`${styles.docType} ${docClass}`}>
            <a href={doc.webUrl} target="_blank">{doc.name}</a>
            <span>{this.utilityMethod.convertDateTime(doc.lastModifiedDateTime, "-")}</span>
          </p>
        );
      });

      const LOBLinks: JSX.Element[] = this.state.LOBLinkList.map(link => (
       

      <div className='col-md-4'>

     
        <div className={`${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4} ${styles.lobApps}`}>
          <a href={link.url} target="_blank" title={link.title} style={{backgroundImage: `url('${link.attachImage}')`}} ></a>
        </div>
        </div>
      ));  

      const myUpcomingEvents: JSX.Element[] = this.state.companyEvents.length>0 ?
      this.state.companyEvents.map(event => {
        return(
        <div className={`${styles.col4Event}`}>
          <div className={`${styles.col4EvtDate}`}>
              <p className={`${styles.col4EvtDay}`}>{event.EventDate.getDate()}</p>
              <p className={`${styles.col4EvtMnth}`}>{this.utilityMethod.monthsOfTheYear[event.EventDate.getMonth()]}</p>
          </div>
          <div className={`${styles.col4EvtDetails}`}>
              <p className={`${styles.col4EvtTitle}`}>{event.Title}</p>
              <p className={`${styles.col4EvtLocatn}`}>{event.Location}</p>
              <p className={`${styles.col4EvtDesc}`} dangerouslySetInnerHTML={{__html: event.Description}}></p>
          </div>
        </div>
        );
      }) :
      [<div className={`${styles.colEmptyContent}`}>
        <p>Hello <strong>{this.state.myProfile ? this.state.myProfile.displayName : "User"}</strong>;<br/> No Organization Event have been booked yet <br/> CHECK BACK LATER!!! </p>
      </div>];

      const myCalendarEvents: JSX.Element[] = this.state.myCalendar.length>0 ?
      this.state.myCalendar.map(event => {
        return (
          <div className={`${styles.col4Event}`}>
            <div>
              <div className={`${styles.col4EvtDate}`}>
                <p className={`${styles.col4EvtDay}`}>
                  {event.startTime.getDate()}
                </p>
                <p className={`${styles.col4EvtMnth}`}>
                  {this.utilityMethod.monthsOfTheYear[event.startTime.getMonth()]}
                </p>
              </div>
              <div className={`${styles.col4EvtDetails}`}>
              <p className={`${styles.col4EvtTitle}`}><a href={event.webLink} target="_blank" >{event.subject}</a></p>
              <p className={`${styles.col4EvtLocatn}`}>
                {
                event.isOnlineMeeting ? <a href={event.onlineMeetingUrl}>TEAMS</a>: event.location.displayName
                  ? event.location.displayName
                  : "Location Unspecified"}
              </p>
              <p
                className={`${styles.col4EvtDesc}`}
                dangerouslySetInnerHTML={{ __html: event.bodyPreview }}
              ></p>
            </div>
            </div>
            {(event.endTime.getTime() - event.startTime.getTime() === 86400000) ?
            <p className={styles.col4EvtTime}>ALL DAY</p> :
            <p className={styles.col4EvtTime}>
              <span>{this.utilityMethod.getFormattedTime(event.startTime)}</span>
              <span>{">>>>>"}</span>
              <span>{this.utilityMethod.getFormattedTime(event.endTime)}</span>
            </p>}
          </div>
        );
      }):
      [<div className={`${styles.colEmptyContent}`}>
        <p>Hello <strong>{this.state.myProfile ? this.state.myProfile.displayName : "User"}</strong>;<br/> You currently don't have any activity scheduled!! Vist <a href="https://outlook.office.com/calendar/?state=0">Outlook</a> to schedule an activity/meeting! </p>
      </div>];

      const LatestHire: JSX.Element[] = [this.state.staffAnniversary.sort((a,z) => z.Resumption_Date.getTime() - a.Resumption_Date.getTime())[0]].map((staff, i) => {
        // if(i<12) 
        return (
        <a className={`ms-Grid-col ms-sm2 ms-md2 ms-lg2 ms-xl2 ms-xxl2 ms-xxxl2 ${styles.msGridCol} ${styles.msSm2}} ${styles.msMd2} ${styles.msLg2} ${styles.msXl2} ${styles.msXxl2} ${styles.msXxxl2} ${styles.col1memeber}`} target="_blank" title={staff&&staff.FullName}>
          <div className={`${styles.col1memberImg}`} style={{backgroundImage: `url('${this.props.siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${staff && staff.Staff_Email}')`}}>
            {staff&&staff.FullName}
          </div>
        </a>
        );
      });
      
  let nextBd: number = 0, currentDate: Date = new Date();
const StaffBirthdays: JSX.Element[] = this.state.staffBirthdays.length>0 ? this.state.staffBirthdays && this.state.staffBirthdays.length>0 && this.state.staffBirthdays.filter(curEmployee => curEmployee.Birthday > (new Date(1900, currentDate.getMonth(), currentDate.getDate())) && curEmployee.Birthday.getDate() >= currentDate.getDate() && curEmployee.Birthday.getMonth() === currentDate.getMonth()).sort((a,b)=>(a.Birthday.getDate()>b.Birthday.getDate())?1:-1).map(birthday => {
  const myBirthday = birthday.Birthday;
  const todayIsBirthday: Boolean = myBirthday.getDate()===currentDate.getDate() && myBirthday.getMonth()===currentDate.getMonth();
  let showNextBd = (nextBd===1 || (!todayIsBirthday && nextBd===0));
 
  nextBd = todayIsBirthday ? 1 : 2;
    
  // console.log(`${myBirthday} - ${todayIsBirthday} - ${showNextBd}`);

  return(
    <div className={`${styles.col4Event} ${todayIsBirthday ? styles.Birtdaybg: ""}`} style={{height:"64px"}}>
      
      <div className={`${styles.col4EvtDate}`}>
          <p className={`${styles.col4EvtDay}`}>{birthday.Birthday.getDate()}</p>
          <p className={`${styles.col4EvtMnth}`}>{this.utilityMethods.monthOfTheYear[birthday["Date_of_Birth"].getMonth()]}</p>
    
      </div>
      <div className={`${styles.col4EvtDetails}`}>
        <p className={styles.BirthdayTitle}
          style={{display: todayIsBirthday ? "block" : "none"}}
        >Happy Birthday To You!</p>
        <p className={`${styles.col4EvtTitle}`}>{birthday["FullName"]}</p>
        <p className={`${styles.col4EvtLocatn}`}>{birthday["Designation"]}</p>
        <p className={`nextBirthdayTitle`}
          style={{display: showNextBd ? "display" : "none"}}
        >Next Birthday</p>
      </div>
    </div>
    
  );
}):
[<div className={`${styles.colEmptyContent}`}>
      <p>Hello <strong>{this.state.myProfile ? this.state.myProfile.displayName : "User"}</strong>;<br/> No Staff Birthday Today <br/> CHECK BACK LATER!!! </p>
    </div>];  

let nextanni: number = 0, currentAnniDate: Date = new Date();
const staffAnniversaryThisMonth = this.state.staffAnniversary && this.state.staffAnniversary.length>0 && this.state.staffAnniversary.filter(curEmployee => curEmployee.Anniversary > (new Date(1900, currentAnniDate.getMonth(), currentAnniDate.getDate())) && curEmployee.Anniversary.getDate() >= currentAnniDate.getDate() && curEmployee.Anniversary.getMonth() === currentAnniDate.getMonth());
const staffAnniversary: JSX.Element[] = staffAnniversaryThisMonth ?
  staffAnniversaryThisMonth.sort((a,b)=>(a.Anniversary.getDate()>b.Anniversary.getDate())?1:-1).map(anniversary => {
  //const myAnniversary = anniversary.anniversaryThisYear;
    const myAnniversary = anniversary.Anniversary;
    // console.log(myAnniversary)
    const todayIsAnniversary: Boolean = myAnniversary.getDate()===currentAnniDate.getDate() && myAnniversary.getMonth()===currentAnniDate.getMonth();
    let showNextAnni = (nextanni===1 || (!todayIsAnniversary && nextanni===0));

    nextBd = todayIsAnniversary ? 1 : 2;
    //console.log(`todayIsAnniversary`, myAnniversary, todayIsAnniversary);

    return(
      <div className={`${styles.col4Event} ${todayIsAnniversary ? styles.Birtdaybg: ""}`}>
        <div className={`${styles.col4EvtDate}`}>
            <p className={`${styles.col4EvtDay}`}>{anniversary.Anniversary.getDate()}</p>
            <p className={`${styles.col4EvtMnth}`}>{this.utilityMethods.monthOfTheYear[anniversary["Resumption_Date"].getMonth()]}</p>
        </div>
        <div className={`${styles.col4EvtDetails}`}>
          <p className={styles.BirthdayTitle}
            style={{display: todayIsAnniversary ? "block" : "none"}}
          >Happy {currentAnniDate.getFullYear() - anniversary.Anniversary.getFullYear()} Year(s) Anniversary To You!</p>
          <p className={`${styles.col4EvtTitle}`}>{anniversary["FullName"]}</p>
          <p className={`${styles.col4EvtLocatn}`}>{anniversary["Designation"]}</p>
          <p className={`nextBirthdayTitle`}
            style={{display: showNextAnni ? "display" : "none"}}
          >Next Anniversary</p>
        </div>
      </div>
    );
  }):
  [<div className={`${styles.colEmptyContent}`}>
        <p>Hello <strong>{this.state.myProfile ? this.state.myProfile.displayName : "User"}</strong>;<br/> No Staff Anniversary Today <br/> CHECK BACK LATER!!! </p>
      </div>];  
      
{/* const StaffBirthday: JSX.Element[] = this.state.staffBirthdays && this.state.staffBirthdays.length>0 ?
      this.state.staffBirthdays.map(bday => {
        return (
          <div className={`${styles.col4Event}`}>
            <div>
              <div className={`${styles.col4EvtDate}`}>
                <p className={`${styles.col4EvtDay}`}>
                  {(new Date(bday["Date_of_Birth"])).getDate()}
                </p>
                <p className={`${styles.col4EvtMnth}`}>
                  {this.utilityMethod.monthsOfTheYear[(new Date(bday["Date_of_Birth"])).getMonth()]}
                </p>
              </div>
              <div className={`${styles.col4EvtDetails}`}>
              <p className={`${styles.col4EvtTitle}`}><a target="_blank" >{bday["First_Name"]}</a></p>
              <p className={`${styles.col4EvtLocatn}`}></p>
              <p
                className={`${styles.col4EvtDesc}`}
              ></p>
            </div>
            </div>
            <p className={styles.col4EvtTime}>ALL DAY</p>
            {/* {(event.endTime.getTime() - event.startTime.getTime() === 86400000) ?
            <p className={styles.col4EvtTime}>ALL DAY</p> :
            <p className={styles.col4EvtTime}>
              <span>{this.utilityMethod.getFormattedTime(event.startTime)}</span>
              <span>{">>>>>"}</span>
              <span>{this.utilityMethod.getFormattedTime(event.endTime)}</span>
            </p>} */}
       {/*   </div>
        );
      }):
      [<div className={`${styles.colEmptyContent}`}>
        <p>Hello <strong>{this.state.myProfile ? this.state.myProfile.displayName : "User"}</strong>;<br/> There's no upcoming birthday on the list</p>
      </div>]; */}
      
    {/*  const staffAnniversary: JSX.Element[] = this.state.staffAnniversary && this.state.staffAnniversary.length>0 ?
      this.state.staffAnniversary.map(Anniday => {
        return (
          <div className={`${styles.col4Event}`}>
            <div>
              <div className={`${styles.col4EvtDate}`}>
                <p className={`${styles.col4EvtDay}`}>
                {(new Date(Anniday["Resumption_Date"])).getDate()}
                </p>
                <p className={`${styles.col4EvtMnth}`}>
                  {this.utilityMethod.monthsOfTheYear[(new Date(Anniday["Resumption_Date"])).getMonth()]}
                </p>
              </div>
              <div className={`${styles.col4EvtDetails}`}>
              <p className={`${styles.col4EvtTitle}`}><a target="_blank" >{Anniday["First_Name"]}</a></p>
              <p className={`${styles.col4EvtLocatn}`}></p>
              <p
                className={`${styles.col4EvtDesc}`}
              ></p> */}
           {/*  </div>
            </div>
            <p className={styles.col4EvtTime}>ALL DAY</p>
            {/*<p className={styles.col4EvtTime}>
              <span>{this.utilityMethod.getFormattedTime(event.startTime)}</span>
              <span>{">>>>>"}</span>
              <span>{this.utilityMethod.getFormattedTime(event.endTime)}</span>
            </p>} 
          </div>
        );
      }):
      [<div className={`${styles.colEmptyContent}`}>
        <p>Hello <strong>{this.state.myProfile ? this.state.myProfile.displayName : "User"}</strong>;<br/> There's no upcoming anniversary on the list</p>
      </div>]; */}

      const employeeSearhResults: JSX.Element[] = this.state.allUser.filter(userProfile => (Boolean(userProfile.mail) && (userProfile.mail.toLowerCase().indexOf(this.state.searchExtn)!=-1 || userProfile.displayName.toLowerCase().indexOf(this.state.searchExtn)!=-1)) /* || userProfile.department.toLowerCase().indexOf(this.state.searchExtn)!=-1 */ ).map((user, i) => {

        return (
          <div className={` ms-Grid-col ms-sm12 ms-md12 ms-lg12 ms-xl12 ms-xxl12 ms-xxxl12 ${styles.msGridCol} ${styles.msSm12}} ${styles.msMd12} ${styles.msLg12} ${styles.msXl12} ${styles.msXxl12} ${styles.msXxxl12} ${styles.employeeSearchDiv} ${styles.employeeSearchDiv2}`}>
            <div>
              <div className={`${ styles.contactCard}`}>
                <div className={`${styles.contactCardHeaderImage}`}>
                  <div className={styles.contactCardAvatar} style={{backgroundImage: `url('${this.props.siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${user.mail}')`}}
                  title={user.displayName}></div>
                </div>
                <div>
                <p className={`${styles.contactCardName} ${styles.contactCardLink} `}>{user.displayName}</p>
                <ul>
                  {user.jobTitle ?
                  <li className={` ${styles.contactCardLinkRole} ${styles.contactCardLink} `}>
                    <p>
                    <span><FontAwesomeIcon icon={'business-time'} ></FontAwesomeIcon>  {user.jobTitle} </span>
                    </p>
                  </li>
                  : ""}
                  {user.department ?
                  <li className={` ${styles.contactCardLinkRole} ${styles.contactCardLink} `}>
                  <p>
                  <span><FontAwesomeIcon icon={'chalkboard-teacher'} ></FontAwesomeIcon>  {user.department}</span>
                  </p>
                  </li>
                  : ""}
                  {user.mail ?
                  <li className={` ${styles.contactCardLinkRole} ${styles.contactCardLink} `}>
                    <p>
                    <span><FontAwesomeIcon icon={'envelope'} ></FontAwesomeIcon>  {user.mail}</span>
                    </p>
                  </li>
                  : ""}
                  {user.mobilePhone || user.businessPhones.join() ?
                  <li className={` ${styles.contactCardLinkExtn} ${styles.contactCardLink} `}>
                    <p>
                    <span><FontAwesomeIcon icon={'mobile-alt'}  ></FontAwesomeIcon> {user.mobilePhone} </span>{user.mobilePhone&&user.businessPhones.join() ? " | " : ""}<span>  {user.businessPhones.join(": ")}</span>
                    </p>
                  </li>
                  : ""}
                </ul>
                </div>
              </div>
            </div>
          </div>
        );
      });

      const spAnnouncementNews: JSX.Element[] = this.state.spAnnouncements.length>0 ?
      this.state.spAnnouncements.map((news) => {
        const createdBy = this.state.spUsers.filter(user => user.Id==news.AuthorId)[0];

        return(
          <div className={styles.newSection}>
            <h5 className={styles.newsTitle}>
              {news.Title}
              <h5 style={{fontSize:"11px", fontWeight: 400}}>{moment.utc(news.Created).local().format('Do MMMM, YYYY')}</h5>
            </h5>
            {/* <h5 className={styles.newsBy}>{`by ${createdBy ? createdBy.Title : `User-${news.AuthorId}`}`}</h5> */}
            <h5 className={styles.newsBy}>{`by ${createdBy ? createdBy.Title : "User-" + news.AuthorId}`}</h5>
            <img className={styles.newsImage} src={news.AttachmentServerURL} width={'90%'} />
            <div className={styles.newsDesc} dangerouslySetInnerHTML={{__html: news.Description}}></div>
            
            <hr className={styles.hr1} />
            <hr className={styles.hr1} />
          </div>
        );
      }):
      [<div className={`${styles.colEmptyContent}`}>
        <p>Hello <strong>{this.state.myProfile ? this.state.myProfile.displayName : "User"}</strong>;<br/> APOLOGIES!!! No Announcement have been made <br/> CHECK BACK LATER!!! </p>
      </div>];

      const plannerTasks: JSX.Element[] = this.state.myPlannerTasks.length>0 ?
        this.state.myPlannerTasks.map((task, i) => {
          const dueDate = new Date(task.dueDateTime),
            dueDateFormatted = `${this.utilityMethod.convertDateTime(dueDate, "-", false, true, true)} ${this.utilityMethod.getFormattedTime(dueDate)}`,
            taskLink = `https://tasks.office.com/${this.state.myProfile.mail.split("@").pop()}/Home/Task/${task.id}?Type=TaskLink`;

          return (
            <div className={styles.plannerTask} onClick={()=>{window.open(taskLink, "_blank");}} style={{cursor: "pointer"}}>
              <div>
                <div className={styles.plannerTaskDetails}>
                  <div>
                    <span>Task</span>
                    <p>{task.title}</p>
                  </div>
                  <div>
                    <span>Due Date</span>
                    <p>{dueDateFormatted}</p>
                  </div>
                </div>
                <div className={styles.plannerTaskStatus}>
                  <p>Status</p>
                  <p>
                    {task.percentComplete === 100
                      ? `Completed`
                      : task.percentComplete === 0
                      ? `Not Started`
                      : `In Progress`}
                  </p>
                  <p>{`${task.percentComplete.toString()}%`}</p>
                </div>
              </div>
              {/* <div>
                <Slider max={100} defaultValue={task.percentComplete} step={5} showValue={true} onChange={value => this.setState({newTaskPercent: value})} />
                <button onClick={() => {
                  this.updateTaskPercent(task.id, task["@odata.etag"], this.state.newTaskPercent);
                }}>Update</button>
                <PrimaryButton text="Update" onClick={() => this.updateTaskPercent(task.id, task["@odata.etag"], this.state.newTaskPercent)} allowDisabledFocus />

              </div> */}
            </div>
          );
        }) :
        [<div className={`${styles.plannerTask} ${styles.emptyPlanner} ${styles.colEmptyContent}`}>
          <p>Hello <strong>{this.state.myProfile ? this.state.myProfile.displayName : "User"}</strong>;<br/> Your outlook planner is empty, visit <a target="_blank" href="https://tasks.office.com/">Office Planner</a> to schedule/assign a new  task! </p>
        </div>];
      
      const eSearchResult: JSX.Element[] = this.state.item.map(searchitem => {
        
        let docClass =  styles.fileDoc0; //className={`${ styles.lipadding}`} extns.indexOf(docExtn) !== -1 ? styles[`${docExtn}Doc0`] :
                
        return(
          <ul className={`ms-List`}>
            <li className={`${ styles.lipadding} ms-ListItem`} >
            {/* <p className={`${docClass}`}></p> */}
            <span className={`${docClass} `}> <i className={`ms-Icon ms-Icon--SharepointAppIcon16`}></i></span>
            <span className={`ms-ListItem-primaryText`}><a href={searchitem.Path} target="_blank">{searchitem.Title}</a></span> 
            {/*  <span className={`ms-ListItem-tertiaryText`}>{searchitem.Description}</span>  */}                  
            </li>
          </ul>
        );
      });

    {/*const Item = styled.div`
    display: flex;
    justify-content: center;
    align-items: center;
    width: calc(100%/2);
    height: 100px;
    font-family: Signika;
    font-weight: bold;
    font-size: 1.5em;
    border: solid 1px black;
    background-color: #61DAFB;
    
    @media(max-width: 1150px) {
        width: 100%
    }
  `;*/}
      
      return (
        // <div className={ `${styles.modernWorkSpace} ${styles.style2} ${styles.esearch} ${this.state.isOnTeams? styles.teamsStyles : ""}` }>
        <div className={ `${styles.modernWorkSpace} ${styles.style2} ${styles.esearch}` } style={{backgroundColor:'White'}}>
        

            {/* test section */}
            <div className={ styles.container }>
          <div className="row">
            {/* slider section */}
            <div className="col-md-12" >
            <div className={` ${styles.column2} ms-sm12 ms-lg12 ${styles.msSm12} ${styles.msLg12} ${styles.colTeams} ${styles.colSPSites} `} style={{borderRadius:'10px',margin:'10px 0px 0px 0px ',backgroundColor:'white',maxWidth:'1088px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px', maxHeight: '327px'}}>
            <div>
                {/* <Carousel
                  buttonsLocation={CarouselButtonsLocation.bottom} buttonsDisplay={CarouselButtonsDisplay.buttonsOnly}
                  element={this.state.currentCarouselItemElement}
                  triggerPageEvent={this.triggerNextElement}
                  contentContainerStyles={{}}
                /> */}
                </div>
              </div>
            </div>

             {/* first Webpart */}
           <div className="col-md-4  ms-sm12 ms-lg4 ">
           <div className={`col-md-4 col-lg-4 col-sm-12 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/bubble-chat.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>Internal Communication</b></div></div>
           <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4}`}  style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px',overflow:'hidden' }}>
                              
                              <div className={`ms-sm12 ms-lg12 ${styles.msSm12} ${styles.msLg12} ${styles.myProfileColumn}`}>
                                <div id="enterpriseSearch" style={{margin: '0 auto 30px'}}>
                                  {/*  <h4 className={styles.column2Title}>E-Search</h4>  */}
                                  <div className={ styles.extnSearch }>
                                    <div className={`${styles.wrap}`}>
                                      <div className={`${styles.search}`} ref={this._menuButtonElement}>
                                        {(this.state.searchText === "")?
                                        <span className={` ${styles.searchSpan}`}>
                                          <i className={`ms-Icon ms-Icon--Search ${styles.searchIconVerticalAlign}`}></i>
                                        </span>
                                        : ""}
                                        <input  name="newregionalcordinatoremailvalue" className={` ${styles.searchTerm}`}  readOnly={false} type="text" placeholder="Search for Content…" value={this.state.searchText}
                                        onChange={(e) => {
                                          this.setState({searchText: e.target["value"]});
                                        }}
                                        onKeyUp={(e) => {
                                          if(e.keyCode===13 || e.key==="Enter"){
                                            this._searchClicked();
                                            // console.log(e.key);
                                          }
                                        }}
                                        style={this.state.searchText===""? {paddingLeft: '15px'}: {}} ></input>
                                        {(this.state.searchText == "")?
                                        ""
                                        :
                                          <div className={` ${styles.searchButtonDiv}`} >
                                            <button type="submit" className={` ${styles.searchClearButton}`} onClick={() => this._searchClearClicked()}>
                                              <i className={`ms-Icon ms-Icon--Clear` }></i>
                                            </button>  
                                            <button type="submit" className={` ${styles.searchButton}`} onClick={() => this._searchClicked()}>
                                              <i className={`ms-Icon ms-Icon--Search ${styles.searchIconVerticalAlign}`} ></i>
                                            </button>
                                          </div>
                                        } 
                                      </div>
              
                                    </div> 
                                  </div> 
                                  { (this.state.searchstatus && this.state.isCalloutVisible) ? 
                                    <Callout
                                      style={{maxWidth: '500px', width: '75%'}}
                                      beakWidth={15}
                                      gapSpace={10}
                                      directionalHint={DirectionalHint.bottomCenter}
                                      target={this._menuButtonElement.current}
                                      onDismiss={(e) => this._onCalloutDismiss(e)}
                                    >
                                      <div className={styles.searchResult}>
                                        <div className={styles.column2Container}>
                                          <div className={styles.docTitle}>
                                          </div>
                                          <div className={styles.docContainer}>
                                            {eSearchResult}
                                          </div>
                                        </div>
                                      </div>
                                    </Callout>
                                    :
                                    <div></div>
                                  }
                                </div>
                              </div>
              
              
                              <div className={`${styles.col1Welcome}`} style={{marginTop: '-61px',overflow:'hidden'}}>
                                <div className={`${styles.col1ProfilePic}`} style={this.state.myProfile != null ? {backgroundImage: `url('${this.props.siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${this.state.myProfile.mail}')`} : {}}></div>
                                <p>WELCOME</p>
                                <p>{this.state.myProfile != null ? this.state.myProfile.displayName : ""}</p>
                                <p><a href={this.state.notificationCount ? "https://outlook.office.com/calendar/view/week" : ""} target="_blank">You have <span>{this.state.notificationCount}</span> notifications</a></p>
                              </div>
                              <div className={`${styles.col1Members}`}>
                                <p>My Recent Contacts</p>
                                <div className={`ms-Grid-row ${styles.msGridRow}`}>{myRecentUser}</div>
                              </div>
                            </div>
           </div>
              {/* second webpart */}
           <div className="col-md-4">
           <div className={`col-md-4 col-lg-4 col-sm-12 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/bubble-chat.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b> My Emails <span>{this.state.mailMessageCount}</span></b></div></div>
           <div className={` ${styles.scrollHidden} ${styles.column} ${styles.colOutlook} ${styles.msLg4} ${styles.zoom} ${styles.msSm12} ${styles.col2}`} style={{borderRadius:'10px',marginTop:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
                  {myMailElArr}
                </div>
           </div>
            {/* third webpart */}
           <div className="col-md-4">
           <div className={`col-md-4 col-lg-4 col-sm-12 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/Teams.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>TEAMS</b></div></div>
           <div className={` ${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} `} style={{borderRadius:'10px',marginTop:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
              <div>
                <h4 className={`${styles.column2Title} ${styles.teamsTitle}`} style={{paddingRight: '55px'}}>Microsoft Teams <img src={require(`./images/setting_gear.svg`)} alt="Teams Centre" title="Teams Centre" width="21px" style={{float: 'right', cursor: 'pointer'}} onClickCapture={() => window.open("https://teams.microsoft.com/_#/apps/bafc60a5-488b-49b6-bc3a-9af2db0a761b/sections/57af5aa1-fef6-43d9-9cc2-a756219cd17f", "_blank")} /></h4>
                <div className={ styles.column2Container }>
                  <div className={styles.colTeamsMyTeams}>
                    {myTeamGroups}
                  </div>
                  <div className={styles.colTeamsConvo}>
                    {myTeamMessages}
                  </div>
                </div>
              </div>
            </div>
          
           </div>
           {/* fourth webpart */}
           {this.props.wbProperties.showCalendar && <div  className="col-md-4">
            <div className={`col-md-4  ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/calendar.svg')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>MY CALENDAR</b></div></div>
            <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '327px', maxHeight: '327',borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1'
            }}>
              {/* <h4>MY CALENDAR</h4> */}
              <div className={`${styles.eventContainer} ${styles.calendarEvent} ${styles.scrollHidden}`}>
                {
                // myUpcomingEvents
                  myCalendarEvents
                }
              </div>
            </div>
            </div>}

            {/* fifth webpart */}
            {this.props.wbProperties.showTasks && <div className="col-md-4">
            <div className={` col-md-4 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/MD Desk.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>Task</b></div></div>
            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} ${styles.colPlanner} `}  style={{borderRadius:'10px',marginTop:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}> 
            <div className={ styles.column2Container }>
                  <div className={` ms-Grid-row ${styles.msGridRow} ${styles.scrollHidden} `} style={{display: this.state.myPlannerTasks.length==0 ? "flex" : "block"}}>
                  {plannerTasks}
                  </div>
                </div>
            
            </div>
            </div>}

          {/* sixth webpart */}
          {this.props.wbProperties.showEvent &&  <div className="col-md-4">
          <div className={`col-md-4 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/EVENT.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>EVENTS</b></div></div>
          <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '327px', maxHeight: '327px',borderRadius:'10px',marginTop:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1'
            }}>
                {/* <h4>EVENTS</h4> */}
                <div className={`${styles.eventContainer} ${styles.scrollHidden}`}>
                  {
                  myUpcomingEvents
                  }
                </div>
              </div>
          </div>}

           {/* Seventh webpart */}
           {this.props.wbProperties.showBirthday && <div  className="col-md-4">
            <div className={`col-md-4  ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/calendar.svg')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>BIRTHDAY</b></div></div>
            <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '327px', maxHeight: '327',borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1'
            }}>
              
               <div className={`${styles.eventContainer} ${styles.calendarEvent} ${styles.scrollHidden}`} style={{display: this.state.staffBirthdays.length==0 ? "flex" : "block"}}>
                {
                 StaffBirthdays
                }
              </div>
            </div>
            </div>}

            {/* Eight webpart */}
            {this.props.wbProperties.showAnnouncement && <div className="col-md-4">
         <div className={`col-md-4 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/Annoucement.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>ANNOUNCEMENT</b></div></div>
         <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} ${styles.colPlanner} `} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
         <div className={ styles.column2Container }>
                  <div className={` ms-Grid-row ${styles.msGridRow} ${styles.scrollHidden} ${styles.announcements} `}>
                  {spAnnouncementNews}
                  </div>
                </div>
            
            </div>
         </div>}

          {/* nineth webpart */}
          {this.props.wbProperties.showAnniversary && <div className="col-md-4">
          <div className={`col-md-4 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/EVENT.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>STAFF ANNIVERSARY</b></div></div>
          <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '327px', maxHeight: '327px',borderRadius:'10px',marginTop:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1'
            }}>
               <div className={`${styles.eventContainer} ${styles.calendarEvent} ${styles.scrollHidden}`} style={{display: this.state.staffAnniversary.length==0 ? "flex" : "block"}}>
                {
                staffAnniversary
                }
              </div>
              </div>
          </div>}


           {/* Tenth webpart */}
           {this.props.wbProperties.showMDDesk && <div className="col-md-4">
          <div className={`col-md-4 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/MD Desk.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>MD's DESK</b></div></div>
          <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} ${styles.colPlanner} `} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
                <div>
                  {/* <h4 className={`${styles.column2Title}`}>Reliance</h4> */}
                 <img src={require('./images/visionmission.png')} style={{
               height : '380px',
               width : '100%'
              }}></img>
    
                </div>
              </div>
          </div>}
            {/* Eleventh webpart */}
            {this.props.wbProperties.showOnedrive && <div className="col-md-8">
          <div className={` col-md-8 ${styles.msSm12} ${styles.msLg12}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',minWidth:'98%',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/microsoft-onedrive-2019.svg')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>ONE DRIVE</b></div></div>
          <div className={` ${styles.column2} ms-sm12 ms-lg8 ${styles.msSm12} ${styles.msLg8} ${styles.colTeams} ${styles.colRecentDoc}  `}style={{borderRadius:'10px',marginTop:'10px',backgroundColor:'white',minWidth:'98%',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
              <div>
                {/* <h4 className={`${styles.column2Title} ${styles.onedriveTitle}`}>OneDrive</h4> */}
                <div className={ styles.column2Container }>
                  <div className={ styles.docTitle}>
                    <span>Name</span><span>Date Modified</span>
                  </div>
                  <div className={ styles.docContainer }>
                    {myRecentDoc}
                  </div>
                </div>
              </div>
            </div>
         </div>}

          {/* first Webpart */}
          {/* <div className="col-md-4  ms-sm12 ms-lg4 ">
           <div className={`col-md-4 col-lg-4 col-sm-12 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/bubble-chat.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>Latest Hire</b></div></div>
           <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4}`}  style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px',overflow:'hidden' }}>
                              
                              <div className={`${styles.col1Welcome}`} style={{marginTop: '60px',overflow:'hidden'}}>
                                <div className={`${styles.col1ProfilePic}`} style={this.state.myProfile != null ? {backgroundImage: `url('${this.props.siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${this.state.staffAnniversary.Staff_Email}')`} : {}}></div>
                                <p>WELCOME</p>
                                <p>{this.state.myProfile != null ? this.state.myProfile.displayName : ""}</p>
                                <p><a href={this.state.notificationCount ? "https://outlook.office.com/calendar/view/week" : ""} target="_blank">You have <span>{this.state.notificationCount}</span> notifications</a></p>
                              </div>
                            </div>
           </div>  */}
           <div className="col-md-4  ms-sm12 ms-lg4 ">
           <div className={`col-md-4 col-lg-4 col-sm-12 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/bubble-chat.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>Latest Hire</b></div></div>
           <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4}`}  style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px',overflow:'hidden' }}>
              <div>
                {/* <h4 className={`${styles.column2Title} ${styles.onedriveTitle}`}>OneDrive</h4> */}
                <div className={ styles.column2Container }>
                {LatestHire}
                </div>
              </div>
            </div>
           </div>

             {/* Eleventh webpart */}
             {this.props.wbProperties.showOnedrive && <div className="col-md-8">
          <div className={` col-md-8 ${styles.msSm12} ${styles.msLg12}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',minWidth:'98%',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/microsoft-onedrive-2019.svg')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>Celebrant of The Month</b></div></div>
          <div className={` ${styles.column2} ms-sm12 ms-lg8 ${styles.msSm12} ${styles.msLg8} ${styles.colTeams} ${styles.colRecentDoc}  `}style={{borderRadius:'10px',marginTop:'10px',backgroundColor:'white',minWidth:'98%',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
              <div>
                {/* <h4 className={`${styles.column2Title} ${styles.onedriveTitle}`}>OneDrive</h4> */}
                <div className={ styles.column2Container }>
                  <div className={ styles.docTitle}>
                    <span>Name</span><span>Date Modified</span>
                  </div>
                  <div className={ styles.docContainer }>
                    {myRecentDoc}
                  </div>
                </div>
              </div>
            </div>
         </div>}

         {/* twelfth webpart */}

         {this.props.wbProperties.showLauncher && <div className="col-md-4">
            <div className={`col-md-4 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/LOB.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>LAUNCHER</b></div></div>
            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} `} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px',overflow:'hidden' }}>
              <div>
                {/* <h4 className={`${styles.column2Title} ${styles.launcherTitle}`}>Launcher</h4> */}
                <div className={ styles.column2Container }>
                  <div className="row">
                    <div className ="col-md-4">
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank" style={{width:"100%"}}>
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--office ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/office_96x1.png)'}}
                      ></div>
                    </a>
                    </div>

                    <div className ="col-md-4">
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/word" target="_blank" style={{width:"100%"}}>
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--word ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/word_96x1.png)'}}
                      ></div>
                    </a>
                    </div>
                    <div className ="col-md-4">
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/excel" target="_blank" style={{width:"100%"}}>
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--excel ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/excel_96x1.png)'}}
                      ></div>
                    </a>
                    </div>
                    <div className ="col-md-4">
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/powerpoint" target="_blank" style={{width:"100%"}}>
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--powerpoint ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/powerpoint_96x1.png)'}}
                      ></div>
                    </a>
                    </div>
                    <div className ="col-md-4">
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/onenote" target="_blank" style={{width:"100%"}}>
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--onenote ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/onenote_96x1.png)'}}
                      ></div>
                    </a>
                    </div>
                    <div className ="col-md-4">
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank"style={{width:"100%"}}>
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--onedrive ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/onedrive_96x1.png)'}}
                      ></div>
                    </a>
                    </div>
                    <div className ="col-md-4">
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank" style={{width:"100%"}}>
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--access ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/access_96x1.png)'}}
                      ></div>
                    </a>
                    </div>
                    <div className ="col-md-4">
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://www.microsoft.com/en-us/microsoft-365/project/project-management-software" target="_blank"style={{width:"100%"}}>
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--project ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/project_96x1.png)'}}
                      ></div>
                    </a>
                    </div>
                    <div className ="col-md-4">
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.live.com/start/visio.aspx" target="_blank" style={{width:"100%"}}>
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--visio ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/visio_96x1.png)'}}
                      ></div>
                    </a>
                    </div>
                  
                  </div>
                 
                </div>
              </div>
            </div>
            </div>}
            {/* thirteenth webpart */}
            {this.props.wbProperties.showLOB && <div className="col-md-4">
            <div className={` col-md-4 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/LOB.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>LINE OF BUSINESS APPS</b></div></div>
            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} `} style={{borderRadius:'10px',marginTop:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
            <div>
            <div className={ styles.column2Container }>
                  <div className="row">
                {/* <h4 className={styles.column2Title}>LINE OF BUSINESS APPS</h4> */}
                {LOBLinks}
                    {/* <div className={`  ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4} ${styles.lobApps}`} style={{width:"100%"}}>
                
                </div> */}
              </div>
            </div>
            </div>
            </div>
            </div>}

            {/* fourteenth webpart */}
            {this.props.wbProperties.showStaffDir && <div className="col-md-4">
            <div className={`col-md-4 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/Employee Analytics.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>STAFF DIRECTORY</b></div></div>
            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} ${styles.colExtns} `}style={{borderRadius:'10px',marginTop:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
              <div>
                {/* <h4 className={styles.column2Title}>STAFF DIRECTORY</h4> */}
                <div className={ styles.column2Container }>
                  <div className={`ms-Grid-row ${styles.msGridRow}`}>
                    <div className={ styles.employeeSearch }>
                      <input type="search" name="extnSearchBox" placeholder="Search for Staff…" id={ styles.employeeSearchBox }
                        onInputCapture={(evt) => this.setState({searchExtn: evt.target["value"].trim().toLowerCase()}) }
                      />
                      <i className="ms-Icon ms-Icon--Search" aria-hidden="true"></i>
                    </div>
                    <div className={` ${styles.extnContainer} ${styles.scrollHidden} `}>
                      {
                      // companyExtns
                      employeeSearhResults
                      }
                    </div>
                  </div>
                </div>
              </div>
            </div>
            </div>}

             {/* thirteenth webpart */}
             {this.props.wbProperties.showDepartmentalAnalytics && <div className="col-md-4">
         <div className={`col-md-4 ${styles.msSm12} ${styles.msLg4}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/Employee Analytics.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>DEPARTMENTAL ANALYTICS</b></div></div>
         <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} ${styles.colSPSites} `} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
              <div>
                {/* <h4 className={`${styles.column2Title} ${styles.analyticsTitle}`}>Departmental Analytics
                  <button className={styles.moreButton}><img src={require('./images/ellipsis1.svg')} width="20px" /></button>
                </h4> */}
                <div className={` ${styles.column2Container} ${styles.scrollHidden} `}>
                  
                  <div className={` ${styles.powerBiContainer} ${styles.analyticsContainer} `} style={{height: 'calc(100% - 30px)'}}>
                    <div style={{width: 'calc(100% - 0px)', display: "inline-block", position: "relative", right: 0}}>
                    {analyticsBarChart2}
                    </div>
                  </div>
                </div>
              </div>
            </div>
        
         </div>}

         {/* fourtheenth webpart */}
         {this.props.wbProperties.showMyAnalytics &&  <div className= "col-md-8">      
         <div className={` col-md-8 ${styles.msSm12} ${styles.msLg8}`} style={{borderRadius:'10px',margin:'10px 0px 0px 0px',backgroundColor:'white',minWidth:'98%',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/analytic.png')} width="6%" height="34px" style={{marginRight:"10px"}}/><b>MY ANALYTICS</b></div></div>
         <div className={` ${styles.column2} ms-sm12 ms-lg8 ${styles.msSm12} ${styles.msLg8} ${styles.colTeams} ${styles.colSPSites} `} style={{borderRadius:'10px',marginTop:'10px',backgroundColor:'white',minWidth:'98%',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px', maxHeight: '327px'}}>
              <div>
                {/* <h4 className={`${styles.column2Title} ${styles.analyticsTitle}`}>MyANALYTICS
                  <button className={styles.moreButton}><img src={require('./images/ellipsis1.svg')} width="20px" /></button>
                </h4> */}
                <div className={` ${styles.column2Container} ${styles.scrollHidden} `} style={{overflowY: "scroll"}}>
                  
                  <div className={` ${styles.analyticsOverview} `}>
                    {Object.keys(this.state.CollaborationActivityList).map(activity => {
                      let duration = this.state.CollaborationActivityList[activity].reduce((acc, cur) => (acc + cur["parsedDuration"]), 0).toFixed(2), trackerColor = activity==="Meeting" ? "#e8c1a0" :  activity==="Focus" ? "#e8a838" :  activity==="Chats/Calls" ? "#f1e15b" :  activity==="Email" ? "#f47560" : "transparent";
                      
                      return (
                      <div className={` ms-sm12 ms-md3 ${styles.msSm12} ${styles.msMd3} ${styles.analyticsCard} `} >
                        <div>
                        <p style={{backgroundImage: `url('${require('./images/analyticsMeeting.svg')}')`}}>{activity}</p>
                        <p><span>{duration}</span><span>hrs</span><span className={`${styles.analyticsCardTracker}`}><span style={{width: `min(calc(100% * ${duration} / 24), 100%)`, backgroundColor: '#e8c1a0'}}></span></span></p>
                        </div>
                      </div>
                      );
                    })}
                  </div>

                  <div className={` ${styles.powerBiContainer} ${styles.analyticsContainer} `}>
                    <div className={` msSm12 msmd6  ${styles.msSm12} ${styles.msMd6}`} style={{display: "inline-block", position: "relative", left: 0}}>
                    {window.matchMedia("(min-width: 480px)").matches ? analyticsPieChart :  analyticsPieChartSM}
                    </div>
                    <div className={` msSm12 msmd6  ${styles.msSm12} ${styles.msMd6}`} style={{display: "inline-block", position: "relative", right: 0}}>
                    {analyticsBarChart1}
                    </div>
                  </div>
                </div>
              </div>
            </div>
            </div>}
         

            </div>
            </div>

              <div className={ styles.row }>
               <div className={`${styles.column2} ms-sm12 ms-lg12 ${styles.msSm8} ${styles.msLg12} ${styles.col4} ${styles.colOpportunities}`} style={{borderRadius:'10px',margin:'10px 10px 10px 40px ',backgroundColor:'white',width:'1030px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'300px', maxHeight: '310px'}}>
                <div>
                {/* <Carousel
                  buttonsLocation={CarouselButtonsLocation.bottom} buttonsDisplay={CarouselButtonsDisplay.buttonsOnly}
                  element={this.state.currentCarouselItemElement}
                  triggerPageEvent={this.triggerNextElement}
                  contentContainerStyles={{}}
                /> */}
                </div>
              </div> 
              </div>

          <div className={ styles.container }>
            <div className='row'>
              <div className={`col-md-4`} style={{borderRadius:'10px',margin:'10px 0px 0px 40px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/bubble-chat.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>Internal Communication</b></div></div>
              <div className={` col-md-4`} style={{borderRadius:'10px',margin:'10px 0px 0px 20px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/bubble-chat.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b> My Emails <span>{this.state.mailMessageCount}</span></b></div></div>
              <div className={`col-md-4`} style={{borderRadius:'10px',margin:'10px 10px 0px 20px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/Teams.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>TEAMS</b></div></div>
              </div>
              <div className={ styles.row }>
              <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4}`}  style={{border:'solid 1px ',borderRadius:'10px',margin:'10px 10px 10px 40px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px',overflow:'hidden' }}>
                              
                <div className={`ms-sm12 ms-lg12 ${styles.msSm12} ${styles.msLg12} ${styles.myProfileColumn}`}>
                  <div id="enterpriseSearch" style={{margin: '0 auto 30px'}}>
                    {/*  <h4 className={styles.column2Title}>E-Search</h4>  */}
                    <div className={ styles.extnSearch }>
                      <div className={`${styles.wrap}`}>
                        <div className={`${styles.search}`} ref={this._menuButtonElement}>
                          {(this.state.searchText === "")?
                          <span className={` ${styles.searchSpan}`}>
                            <i className={`ms-Icon ms-Icon--Search ${styles.searchIconVerticalAlign}`}></i>
                          </span>
                          : ""}
                          <input  name="newregionalcordinatoremailvalue" className={` ${styles.searchTerm}`}  readOnly={false} type="text" placeholder="Search for Content…" value={this.state.searchText}
                          onChange={(e) => {
                            this.setState({searchText: e.target["value"]});
                          }}
                          onKeyUp={(e) => {
                            if(e.keyCode===13 || e.key==="Enter"){
                              this._searchClicked();
                              // console.log(e.key);
                            }
                          }}
                          style={this.state.searchText===""? {paddingLeft: '15px'}: {}} ></input>
                          {(this.state.searchText == "")?
                          ""
                          :
                            <div className={` ${styles.searchButtonDiv}`} >
                              <button type="submit" className={` ${styles.searchClearButton}`} onClick={() => this._searchClearClicked()}>
                                <i className={`ms-Icon ms-Icon--Clear` }></i>
                              </button>  
                              <button type="submit" className={` ${styles.searchButton}`} onClick={() => this._searchClicked()}>
                                <i className={`ms-Icon ms-Icon--Search ${styles.searchIconVerticalAlign}`} ></i>
                              </button>
                            </div>
                          } 
                        </div>

                      </div> 
                    </div> 
                    { (this.state.searchstatus && this.state.isCalloutVisible) ? 
                      <Callout
                        style={{maxWidth: '500px', width: '75%'}}
                        beakWidth={15}
                        gapSpace={10}
                        directionalHint={DirectionalHint.bottomCenter}
                        target={this._menuButtonElement.current}
                        onDismiss={(e) => this._onCalloutDismiss(e)}
                      >
                        <div className={styles.searchResult}>
                          <div className={styles.column2Container}>
                            <div className={styles.docTitle}>
                            </div>
                            <div className={styles.docContainer}>
                              {eSearchResult}
                            </div>
                          </div>
                        </div>
                      </Callout>
                      :
                      <div></div>
                    }
                  </div>
                </div>


                <div className={`${styles.col1Welcome}`} style={{marginTop: '-61px',overflow:'hidden'}}>
                  <div className={`${styles.col1ProfilePic}`} style={this.state.myProfile != null ? {backgroundImage: `url('${this.props.siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${this.state.myProfile.mail}')`} : {}}></div>
                  <p>WELCOME</p>
                  <p>{this.state.myProfile != null ? this.state.myProfile.displayName : ""}</p>
                  <p><a href={this.state.notificationCount ? "https://outlook.office.com/calendar/view/week" : ""} target="_blank">You have <span>{this.state.notificationCount}</span> notifications</a></p>
                </div>
                <div className={`${styles.col1Members}`}>
                  <p>My Recent Contacts</p>
                  <div className={`ms-Grid-row ${styles.msGridRow}`}>{myRecentUser}</div>
                </div>
              </div>
              <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.zoom} ${styles.msSm12} ${styles.msLg4} ${styles.col2} ${styles.colOutlook}`} style={{borderRadius:'10px',margin:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
                {/* <p className={`${styles.col2Notification}`}>MY EMAILS <span>{this.state.mailMessageCount}</span></p> */}
                <div className={` ${styles.scrollHidden}`}>
                  {myMailElArr}
                </div>
              </div>
              <div className={` ${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} `} style={{border:'solid 1px',borderRadius:'10px',margin:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
              <div>
                <h4 className={`${styles.column2Title} ${styles.teamsTitle}`} style={{paddingRight: '55px'}}>Microsoft Teams <img src={require(`./images/setting_gear.svg`)} alt="Teams Centre" title="Teams Centre" width="21px" style={{float: 'right', cursor: 'pointer'}} onClickCapture={() => window.open("https://teams.microsoft.com/_#/apps/bafc60a5-488b-49b6-bc3a-9af2db0a761b/sections/57af5aa1-fef6-43d9-9cc2-a756219cd17f", "_blank")} /></h4>
                <div className={ styles.column2Container }>
                  <div className={styles.colTeamsMyTeams}>
                    {myTeamGroups}
                  </div>
                  <div className={styles.colTeamsConvo}>
                    {myTeamMessages}
                  </div>
                </div>
              </div>
            </div>
            </div>
           
          </div>
          <div className={ styles.container }>
          <div className='row'>
          {this.props.wbProperties.showCalendar && <div className={`col-md-4`} style={{borderRadius:'10px',margin:'10px 0px 0px 40px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/calendar.svg')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>MY CALENDAR</b></div></div>}
              {this.props.wbProperties.showTasks && <div className={` col-md-4`} style={{borderRadius:'10px',margin:'10px 0px 0px 20px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/MD Desk.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>MY TASK</b></div></div>}
              <div className={`col-md-4`} style={{borderRadius:'10px',margin:'10px 10px 0px 20px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/EVENT.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>EVENTS</b></div></div>
              </div>
            <div className={ styles.row }>
            {this.props.wbProperties.showCalendar && <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '327px', maxHeight: '327',border:'solid 1px ',borderRadius:'10px',margin:'10px 10px 10px 40px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1'
            }}>
              {/* <h4>MY CALENDAR</h4> */}
              <div className={`${styles.eventContainer} ${styles.calendarEvent} ${styles.scrollHidden}`}>
                {
                // myUpcomingEvents
                  myCalendarEvents
                }
              </div>
            </div> }
            
            {this.props.wbProperties.showTasks && <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} ${styles.colPlanner} `}  style={{borderRadius:'10px',margin:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}> 
                {/* <h4 className={styles.column2Title}>ANNOUNCEMENTS</h4> */}
                <div className={ styles.column2Container }>
                  <div className={` ms-Grid-row ${styles.msGridRow} ${styles.scrollHidden} `} style={{display: this.state.myPlannerTasks.length==0 ? "flex" : "block"}}>
                  {plannerTasks}
                  </div>
                </div>
            
            </div>}
            <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '327px', maxHeight: '327px',border:'solid 1px',borderRadius:'10px',margin:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1'
            }}>
                {/* <h4>EVENTS</h4> */}
                <div className={`${styles.eventContainer} ${styles.scrollHidden}`}>
                  {
                  myUpcomingEvents
                  }
                </div>
              </div>
              </div> 

              <div className='row'>
              <div className={`col-md-4`} style={{borderRadius:'10px',margin:'10px 0px 0px 40px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/calendar.svg')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>STAFF BIRTHDAY'S</b></div></div>
              <div className={` col-md-4`} style={{borderRadius:'10px',margin:'10px 0px 0px 20px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/Annoucement.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>ANNOUNCEMENT</b></div></div>
              <div className={`col-md-4`} style={{borderRadius:'10px',margin:'10px 10px 0px 20px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/EVENT.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>STAFF ANNIVERSARY</b></div></div>
              </div>
            <div className={ styles.row }>
            <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '327px', maxHeight: '327',border:'solid 1px ',borderRadius:'10px',margin:'10px 10px 10px 40px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1'
            }}>
              {/* <h4>MY CALENDAR</h4> */}
              <div className={`${styles.eventContainer} ${styles.calendarEvent} ${styles.scrollHidden}`} style={{display: this.state.staffBirthdays.length==0 ? "flex" : "block"}}>
                {
                 StaffBirthdays
                }
              </div>
            </div>
            
            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} ${styles.colPlanner} `}  style={{borderRadius:'10px',margin:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}> 
                {/* <h4 className={styles.column2Title}>ANNOUNCEMENTS</h4> */}
                <div className={ styles.column2Container }>
                  <div className={` ms-Grid-row ${styles.msGridRow} ${styles.scrollHidden} ${styles.announcements} `}>
                  {spAnnouncementNews}
                  </div>
                </div>
            
            </div>
            <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '327px', maxHeight: '327px',border:'solid 1px',borderRadius:'10px',margin:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1'
            }}>
                {/* <h4>EVENTS</h4> */}
                <div className={`${styles.eventContainer} ${styles.calendarEvent} ${styles.scrollHidden}`} style={{display: this.state.staffAnniversary.length==0 ? "flex" : "block"}}>
                {
                staffAnniversary
                }
              </div>
              </div>
 
              </div> 


              
              <div className='row'>
              <div className={`col-md-4`} style={{borderRadius:'10px',margin:'10px 0px 0px 40px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/MD Desk.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>MD's DESK</b></div></div>
              <div className={` col-md-8`} style={{borderRadius:'10px',margin:'10px 0px 0px 20px',backgroundColor:'white',maxWidth:'680px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/microsoft-onedrive-2019.svg')} width="7%" height="34px" style={{marginRight:"10px"}}/><b>ONE DRIVE</b></div></div>
            
              </div>
              <div className={ styles.row }>
              <div className={`${styles.column2} ms-sm12 ms-lg4  ${styles.msLg4} ${styles.col4} ${styles.colOpportunities}`} style={{
               height: '327px', maxHeight: '327',border:'solid 1px ',borderRadius:'10px',margin:'10px 10px 10px 40px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1'
              }}>
                <div>
                  {/* <h4 className={`${styles.column2Title}`}>Reliance</h4> */}
                 <img src={require('./images/visionmission.png')} style={{
               height : '380px',
               width : '400px'
              }}></img>
    
                </div>
              </div>
            <div className={` ${styles.column2} ms-sm12 ms-lg8 ${styles.msSm12} ${styles.msLg8} ${styles.colTeams} ${styles.colRecentDoc}  `}style={{borderRadius:'10px',margin:'10px',backgroundColor:'white',maxWidth:'680px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
              <div>
                {/* <h4 className={`${styles.column2Title} ${styles.onedriveTitle}`}>OneDrive</h4> */}
                <div className={ styles.column2Container }>
                  <div className={ styles.docTitle}>
                    <span>Name</span><span>Date Modified</span>
                  </div>
                  <div className={ styles.docContainer }>
                    {myRecentDoc}
                  </div>
                </div>
              </div>
            </div>
              </div>
            
            </div> 
            

 
            <div className='row'>
              <div className={`col-md-4`} style={{borderRadius:'10px',margin:'10px 0px 0px 40px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/LOB.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>LAUNCHER</b></div></div>
              <div className={` col-md-4`} style={{borderRadius:'10px',margin:'10px 0px 0px 20px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/LOB.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>LINE OF BUSINESS APPS</b></div></div>
              <div className={`col-md-4`} style={{borderRadius:'10px',margin:'10px 10px 0px 20px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/Employee Analytics.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>STAFF DIRECTORY</b></div></div>
              </div>

              <div className={ styles.row }>
            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} `} style={{border:'solid 1px ',borderRadius:'10px',margin:'10px 10px 10px 40px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px',overflow:'hidden' }}>
              <div>
                {/* <h4 className={`${styles.column2Title} ${styles.launcherTitle}`}>Launcher</h4> */}
                <div className={ styles.column2Container }>
                  <div className={` ms-Grid-row ${styles.msGridRow} ${styles.scrollHidden} `}>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--office ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/office_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/word" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--word ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/word_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/excel" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--excel ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/excel_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/powerpoint" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--powerpoint ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/powerpoint_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--access ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/access_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/onenote" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--onenote ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/onenote_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--onedrive ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/onedrive_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://www.microsoft.com/en-us/microsoft-365/project/project-management-software" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--project ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/project_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.live.com/start/visio.aspx" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--visio ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/visio_96x1.png)'}}
                      ></div>
                    </a>
                  </div>
                </div>
              </div>
            </div>

            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} `} style={{borderRadius:'10px',margin:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
              <div>
                {/* <h4 className={styles.column2Title}>LINE OF BUSINESS APPS</h4> */}
                <div className={ styles.column2Container }>
                  <div className={` ms-Grid-row ${styles.msGridRow} `} style={{ overflowY: "scroll" }}>
                    {LOBLinks}
                  </div>
                </div>
              </div>
            </div>
            
            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} ${styles.colExtns} `}style={{borderRadius:'10px',margin:'10px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
              <div>
                {/* <h4 className={styles.column2Title}>STAFF DIRECTORY</h4> */}
                <div className={ styles.column2Container }>
                  <div className={`ms-Grid-row ${styles.msGridRow}`}>
                    <div className={ styles.employeeSearch }>
                      <input type="search" name="extnSearchBox" placeholder="Search for Staff…" id={ styles.employeeSearchBox }
                        onInputCapture={(evt) => this.setState({searchExtn: evt.target["value"].trim().toLowerCase()}) }
                      />
                      <i className="ms-Icon ms-Icon--Search" aria-hidden="true"></i>
                    </div>
                    <div className={` ${styles.extnContainer} ${styles.scrollHidden} `}>
                      {
                      // companyExtns
                      employeeSearhResults
                      }
                    </div>
                  </div>
                </div>
              </div>
            </div>
            </div>

            <div className='row'>
              <div className={`col-md-4`} style={{borderRadius:'10px',margin:'10px 0px 0px 42px',backgroundColor:'white',maxWidth:'330px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/Employee Analytics.png')} width="13%" height="34px" style={{marginRight:"10px"}}/><b>DEPARTMENTAL ANALYTICS</b></div></div>
              <div className={` col-md-8`} style={{borderRadius:'10px',margin:'10px 0px 0px 20px',backgroundColor:'white',maxWidth:'680px',height:'53px',color:'#1e90ff',paddingTop:'10px',boxShadow : '5px 4px 9px 10px #f1f1f1'}}><div style={{flexDirection:"row",justifyContent:"space-between"}}><img src={require('./images/analytic.png')} width="6%" height="34px" style={{marginRight:"10px"}}/><b>MY ANALYTICS</b></div></div>
            
              </div>

          
              <div className={ styles.row }>
            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} ${styles.colSPSites} `} style={{borderRadius:'10px',margin:'10px 10px 10px 38px',backgroundColor:'white',maxWidth:'330px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px'}}>
              <div>
                {/* <h4 className={`${styles.column2Title} ${styles.analyticsTitle}`}>Departmental Analytics
                  <button className={styles.moreButton}><img src={require('./images/ellipsis1.svg')} width="20px" /></button>
                </h4> */}
                <div className={` ${styles.column2Container} ${styles.scrollHidden} `}>
                  
                  <div className={` ${styles.powerBiContainer} ${styles.analyticsContainer} `} style={{height: 'calc(100% - 30px)'}}>
                    <div style={{width: 'calc(100% - 0px)', display: "inline-block", position: "relative", right: 0}}>
                    {analyticsBarChart1}
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div className={` ${styles.column2} ms-sm12 ms-lg8 ${styles.msSm12} ${styles.msLg8} ${styles.colTeams} ${styles.colSPSites} `} style={{borderRadius:'10px',margin:'10px',backgroundColor:'white',width:'680px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px', maxHeight: '327px'}}>
              <div>
                {/* <h4 className={`${styles.column2Title} ${styles.analyticsTitle}`}>MyANALYTICS
                  <button className={styles.moreButton}><img src={require('./images/ellipsis1.svg')} width="20px" /></button>
                </h4> */}
                <div className={` ${styles.column2Container} ${styles.scrollHidden} `} style={{overflowY: "scroll"}}>
                  
                  <div className={` ${styles.analyticsOverview} `}>
                    {Object.keys(this.state.CollaborationActivityList).map(activity => {
                      let duration = this.state.CollaborationActivityList[activity].reduce((acc, cur) => (acc + cur["parsedDuration"]), 0).toFixed(2), trackerColor = activity==="Meeting" ? "#e8c1a0" :  activity==="Focus" ? "#e8a838" :  activity==="Chats/Calls" ? "#f1e15b" :  activity==="Email" ? "#f47560" : "transparent";
                      
                      return (
                      <div className={` ms-sm12 ms-md3 ${styles.msSm12} ${styles.msMd3} ${styles.analyticsCard} `} >
                        <div>
                        <p style={{backgroundImage: `url('${require('./images/analyticsMeeting.svg')}')`}}>{activity}</p>
                        <p><span>{duration}</span><span>hrs</span><span className={`${styles.analyticsCardTracker}`}><span style={{width: `min(calc(100% * ${duration} / 24), 100%)`, backgroundColor: '#e8c1a0'}}></span></span></p>
                        </div>
                      </div>
                      );
                    })}
                  </div>

                  <div className={` ${styles.powerBiContainer} ${styles.analyticsContainer} `}>
                    <div className={` msSm12 msmd6  ${styles.msSm12} ${styles.msMd6}`} style={{display: "inline-block", position: "relative", left: 0}}>
                    {window.matchMedia("(min-width: 480px)").matches ? analyticsPieChart :  analyticsPieChartSM}
                    </div>
                    <div className={` msSm12 msmd6  ${styles.msSm12} ${styles.msMd6}`} style={{display: "inline-block", position: "relative", right: 0}}>
                    {analyticsBarChart1}
                    </div>
                  </div>
                </div>
              </div>
            </div>
           </div>







       {/*  
          <div className={ styles.row }>
               <div className={`${styles.column2} ms-sm12 ms-lg12 ${styles.msSm8} ${styles.msLg12} ${styles.col4} ${styles.colOpportunities}`} style={{borderRadius:'10px',margin:'10px 10px 10px 30px ',backgroundColor:'white',width:'1040px',boxShadow : '5px 4px 9px 10px #f1f1f1',height:'327px', maxHeight: '327px'}}>
                <div>
                   <Carousel
                  buttonsLocation={CarouselButtonsLocation.bottom} buttonsDisplay={CarouselButtonsDisplay.buttonsOnly}
                  element={this.state.currentCarouselItemElement}
                  triggerPageEvent={this.triggerNextElement}
                  contentContainerStyles={{}}
                />
                </div>
              </div> 
              </div> */}

         {/*  <div className={ `${styles.container} ${styles.carouselContainer}` }>
            <div className={ styles.row }>
              <div className={`${styles.column} ms-sm12 ms-lg12 ${styles.msSm12} ${styles.msLg12}`} style={{
                height: '340px', maxHeight: '340px', marginBottom: '20px'
              }}>
                 {/* <Carousel showThumbs={true} autoPlay={true} infiniteLoop={true} interval={5000} autoFocus={true} transitionTime={5000} >
                  {this.state.carouselItems.map((imageList) => {
                    return (<div style={{display:"flex", height:"100%"}}>
                     <a href={imageList.ResourceLink} target='_blank'> <img src={imageList.FileRef} /> </a> 
                    </div>);
                  })}    
                </Carousel>  */}
              {/*   <Carousel
                  buttonsLocation={CarouselButtonsLocation.bottom} buttonsDisplay={CarouselButtonsDisplay.buttonsOnly}
                  element={this.state.currentCarouselItemElement}
                  triggerPageEvent={this.triggerNextElement}
                  contentContainerStyles={{}}
                />
               {/* <MyCarousel duration={5000}>
                  {this.state.carouselItems.map((imageList) => {
                    return (<MyCarouselItem width='100%'>
                      <div style={{display:"flex", height:"100%"}}>
                        <a href={imageList.ResourceLink} target='_blank'> <img src={imageList.FileRef} /> </a> 
                      </div>
                      </MyCarouselItem>);
                  })}
                </MyCarousel> */}
               {/* <Rerousel itemRef={ref}>
      
                </Rerousel> */}

             {/*  </div>
            </div>
          </div>*/}

         {/*  <div className={ `${styles.container} ${styles.firstContainer}` }>
            <div className={ styles.row }>

              <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4}`}>
                              
                <div className={`ms-sm12 ms-lg12 ${styles.msSm12} ${styles.msLg12} ${styles.myProfileColumn}`}>
                  <div id="enterpriseSearch" style={{margin: '0 auto 30px'}}>
                    {/*  <h4 className={styles.column2Title}>E-Search</h4>  */}
                   {/*  <div className={ styles.extnSearch }>
                      <div className={`${styles.wrap}`}>
                        <div className={`${styles.search}`} ref={this._menuButtonElement}>
                          {(this.state.searchText === "")?
                          <span className={` ${styles.searchSpan}`}>
                            <i className={`ms-Icon ms-Icon--Search ${styles.searchIconVerticalAlign}`}></i>
                          </span>
                          : ""}
                          <input  name="newregionalcordinatoremailvalue" className={` ${styles.searchTerm}`}  readOnly={false} type="text" placeholder="Search for Content…" value={this.state.searchText}
                          onChange={(e) => {
                            this.setState({searchText: e.target["value"]});
                          }}
                          onKeyUp={(e) => {
                            if(e.keyCode===13 || e.key==="Enter"){
                              this._searchClicked();
                              // console.log(e.key);
                            }
                          }}
                          style={this.state.searchText===""? {paddingLeft: '15px'}: {}} ></input>
                          {(this.state.searchText == "")?
                          ""
                          :
                            <div className={` ${styles.searchButtonDiv}`} >
                              <button type="submit" className={` ${styles.searchClearButton}`} onClick={() => this._searchClearClicked()}>
                                <i className={`ms-Icon ms-Icon--Clear` }></i>
                              </button>  
                              <button type="submit" className={` ${styles.searchButton}`} onClick={() => this._searchClicked()}>
                                <i className={`ms-Icon ms-Icon--Search ${styles.searchIconVerticalAlign}`} ></i>
                              </button>
                            </div>
                          } 
                        </div>

                      </div> 
                    </div> 
                    { (this.state.searchstatus && this.state.isCalloutVisible) ? 
                      <Callout
                        style={{maxWidth: '500px', width: '75%'}}
                        beakWidth={15}
                        gapSpace={10}
                        directionalHint={DirectionalHint.bottomCenter}
                        target={this._menuButtonElement.current}
                        onDismiss={(e) => this._onCalloutDismiss(e)}
                      >
                        <div className={styles.searchResult}>
                          <div className={styles.column2Container}>
                            <div className={styles.docTitle}>
                            </div>
                            <div className={styles.docContainer}>
                              {eSearchResult}
                            </div>
                          </div>
                        </div>
                      </Callout>
                      :
                      <div></div>
                    }
                  </div>
                </div>


                <div className={`${styles.col1Welcome}`}>
                  <div className={`${styles.col1ProfilePic}`} style={this.state.myProfile != null ? {backgroundImage: `url('${this.props.siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${this.state.myProfile.mail}')`} : {}}></div>
                  <p>WELCOME</p>
                  <p>{this.state.myProfile != null ? this.state.myProfile.displayName : ""}</p>
                  <p><a href={this.state.notificationCount ? "https://outlook.office.com/calendar/view/week" : "https://outlook.office.com/calendar/view/week"} target="_blank">You have <span>{this.state.notificationCount}</span> notifications</a></p>
                </div>
                <div className={`${styles.col1Members}`}>
                  <p>My Recent Contacts</p>
                  <div className={`ms-Grid-row ${styles.msGridRow}`}>{myRecentUser}</div>
                </div>
              </div>
              
              <div className={`${styles.column} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col2} ${styles.colOutlook}`}>
                <p className={`${styles.col2Notification}`}>My Emails<span>{this.state.mailMessageCount}</span></p>
                <div className={` ${styles.outlookMailContainer} ${styles.scrollHidden}`}>
                  {myMailElArr}
                </div>
              </div>
              
              <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} `}>
                <div>
                  <h4 className={`${styles.column2Title} ${styles.teamsTitle}`} style={{paddingRight: '55px'}}>Microsoft Teams <img src={require(`./images/setting_gear.svg`)} alt="Teams Centre" title="Teams Centre" width="21px" style={{float: 'right', cursor: 'pointer'}} onClickCapture={() => window.open("https://teams.microsoft.com/_#/apps/bafc60a5-488b-49b6-bc3a-9af2db0a761b/sections/57af5aa1-fef6-43d9-9cc2-a756219cd17f", "_blank")} /></h4>
                  <div className={ styles.column2Container }>
                    <div className={styles.colTeamsMyTeams}>
                      {myTeamGroups}
                    </div>
                    <div className={styles.colTeamsConvo}>
                      {myTeamMessages}
                    </div>
                  </div>
                </div>
              </div>
            
            </div>
          </div> */}

         {/*  <div className={ styles.container }>
          <div className={ styles.row }>
              
            {this.props.wbProperties.showTasks && <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} ${styles.colPlanner} `}>
              <div>
                <h4 className={`${styles.column2Title} ${styles.taskTitle}`}>My Tasks</h4>
                <div className={ styles.column2Container }>
                  <div className={` ms-Grid-row ${styles.msGridRow} ${styles.scrollHidden} `} style={{display: this.state.myPlannerTasks.length==0 ? "flex" : "block"}}>
                  {plannerTasks}
                  </div>
                </div>
              </div>
            </div>}

            {this.props.wbProperties.showCalendar && <div className={`${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '400px', maxHeight: '400px'
              }}>
              <div>
                <h4 className={`${styles.column2Title}`}>My Calendar</h4>
                <div className={`${styles.eventContainer} ${styles.calendarEvent} ${styles.scrollHidden}`} style={{display: this.state.myCalendar.length==0 ? "flex" : "block"}}>
                  {
                  // myUpcomingEvents
                    myCalendarEvents
                  }
                </div>
              </div>
            </div>}

            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} `}>
              <div>
                <h4 className={styles.column2Title}>SharePoint Business Process Requests</h4>
                <div className={ styles.column2Container }>
                  <div className={` ms-Grid-row ${styles.msGridRow} `} style={{ overflowY: "scroll" }}>
                    {LOBLinks}
                  </div>
                </div>
              </div>
            </div> */}

           {/*  <div className={`${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '400px', maxHeight: '400px'
            }}>
              <div>
              <h4 className={`${styles.column2Title}`}>Staff Birthdays</h4>
              <div className={`${styles.eventContainer} ${styles.calendarEvent} ${styles.scrollHidden}`} style={{display: this.state.staffBirthdays.length==0 ? "flex" : "block"}}>
                {
                 StaffBirthdays
                }
              </div>
              </div>
            </div>

            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} ${styles.colPlanner} `}>
              <div>
                <h4 className={styles.column2Title}>Announcements</h4>
                <div className={ styles.column2Container }>
                  <div className={` ms-Grid-row ${styles.msGridRow} ${styles.scrollHidden} ${styles.announcements} `}>
                  {spAnnouncementNews}
                  </div>
                </div>
              </div>
            </div> */}
            
             {/*  <div className={`${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '400px', maxHeight: '400px'
            }}>
              <div>
              <h4 className={`${styles.column2Title}`}>Company Events</h4>
              <div className={`${styles.eventContainer} ${styles.scrollHidden}`} style={{display: this.state.companyEvents.length==0 ? "flex" : "block"}}>
                {
                myUpcomingEvents
                }
              </div>
              </div>
            </div> */}
            

           {/*  <div className={`${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.col4}`} style={{
              height: '400px', maxHeight: '400px'
            }}>
              <div>
              <h4 className={`${styles.column2Title}`}>Anniversary</h4>
              <div className={`${styles.eventContainer} ${styles.calendarEvent} ${styles.scrollHidden}`} style={{display: this.state.staffAnniversary.length==0 ? "flex" : "block"}}>
                {
                staffAnniversary
                }
              </div>
              </div>
            </div>
            

            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} `}>
              <div>
                <h4 className={`${styles.column2Title} ${styles.launcherTitle}`}>Launcher</h4>
                <div className={ styles.column2Container }>
                  <div className={` ms-Grid-row ${styles.msGridRow} ${styles.scrollHidden} `}>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--office ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/office_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/word" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--word ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/word_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/excel" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--excel ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/excel_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/powerpoint" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--powerpoint ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/powerpoint_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--access ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/access_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/onenote" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--onenote ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/onenote_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--onedrive ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/onedrive_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://www.microsoft.com/en-us/microsoft-365/project/project-management-software" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--project ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/project_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.live.com/start/visio.aspx" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--visio ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/visio_96x1.png)'}}
                      ></div>
                    </a>
                  </div>
                </div>
              </div>
            </div> */}

            {/* <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} ${styles.colRecentDoc}  `}>
              <div>
                <h4 className={`${styles.column2Title} ${styles.onedriveTitle}`}>OneDrive</h4>
                <div className={ styles.column2Container }>
                  <div className={ styles.docTitle}>
                    <span>Name</span><span>Date Modified</span>
                  </div>
                  <div className={ styles.docContainer }>
                    {myRecentDoc}
                  </div>
                </div>
              </div>
            </div>
            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} ${styles.colExtns} `}>
              <div>
                <h4 className={styles.column2Title}>Staff Directory</h4>
                <div className={ styles.column2Container }>
                  <div className={`ms-Grid-row ${styles.msGridRow}`}>
                    <div className={ styles.employeeSearch }>
                      <input type="search" name="extnSearchBox" placeholder="Search for Staff…" id={ styles.employeeSearchBox }
                        onInputCapture={(evt) => this.setState({searchExtn: evt.target["value"].trim().toLowerCase()}) }
                      />
                      <i className="ms-Icon ms-Icon--Search" aria-hidden="true"></i>
                    </div>
                    <div className={` ${styles.extnContainer} ${styles.scrollHidden} `}>
                      {
                      // companyExtns
                      employeeSearhResults
                      }
                    </div>
                  </div>
                </div>
              </div>
            </div>


            <div className={` ${styles.column2} ms-sm12 ms-lg12 ${styles.msSm12} ${styles.msLg12} ${styles.colTeams} ${styles.colSPSites} `} style={{
              height: '450px', maxHeight: '450px'
            }}>
              <div>
                <h4 className={`${styles.column2Title} ${styles.analyticsTitle}`}>MyANALYTICS
                  <button className={styles.moreButton}><img src={require('./images/ellipsis1.svg')} width="20px" /></button>
                </h4>
                <div className={` ${styles.column2Container} ${styles.scrollHidden} `} style={{overflowY: "scroll"}}>
                  
                  <div className={` ${styles.analyticsOverview} `}>
                    {Object.keys(this.state.CollaborationActivityList).map(activity => {
                      let duration = this.state.CollaborationActivityList[activity].reduce((acc, cur) => (acc + cur["parsedDuration"]), 0).toFixed(2), trackerColor = activity==="Meeting" ? "#e8c1a0" :  activity==="Focus" ? "#e8a838" :  activity==="Chats/Calls" ? "#f1e15b" :  activity==="Email" ? "#f47560" : "transparent";
                      
                      return (
                      <div className={` ms-sm12 ms-md3 ${styles.msSm12} ${styles.msMd3} ${styles.analyticsCard} `} >
                        <div>
                        <p style={{backgroundImage: `url('${require('./images/analyticsMeeting.svg')}')`}}>{activity}</p>
                        <p><span>{duration}</span><span>hrs</span><span className={`${styles.analyticsCardTracker}`}><span style={{width: `min(calc(100% * ${duration} / 24), 100%)`, backgroundColor: '#e8c1a0'}}></span></span></p>
                        </div>
                      </div>
                      );
                    })}
                  </div>

                  <div className={` ${styles.powerBiContainer} ${styles.analyticsContainer} `}>
                    <div className={` msSm12 msmd6  ${styles.msSm12} ${styles.msMd6}`} style={{display: "inline-block", position: "relative", left: 0}}>
                    {window.matchMedia("(min-width: 480px)").matches ? analyticsPieChart :  analyticsPieChartSM}
                    </div>
                    <div className={` msSm12 msmd6  ${styles.msSm12} ${styles.msMd6}`} style={{display: "inline-block", position: "relative", right: 0}}>
                    {analyticsBarChart1}
                    </div>
                  </div>
                </div>
              </div>
            </div> */}


            {/* 
            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} `}>
              <div>
                <h4 className={`${styles.column2Title} ${styles.teamsTitle}`} style={{paddingRight: '55px'}}>Microsoft Teams <img src={require(`./images/setting_gear.svg`)} alt="Teams Centre" title="Teams Centre" width="21px" style={{float: 'right', cursor: 'pointer'}} onClickCapture={() => window.open("https://teams.microsoft.com/_#/apps/bafc60a5-488b-49b6-bc3a-9af2db0a761b/sections/57af5aa1-fef6-43d9-9cc2-a756219cd17f", "_blank")} /></h4>
                <div className={ styles.column2Container }>
                  <div className={styles.colTeamsMyTeams}>
                    {myTeamGroups}
                  </div>
                  <div className={styles.colTeamsConvo}>
                    {myTeamMessages}
                  </div>
                </div>
              </div>
            </div>
            <div className={` ${styles.column2} ms-sm12 ms-lg4  ${styles.msSm12} ${styles.msLg8} ${styles.colLauncher} ${styles.colExtns} `}>
              <div>
              <h4 className={styles.column2Title}>ENGAGEMENT PROGRESS</h4>
              <div className={ styles.column2Container }>
                <div className="ms-Grid-row">
                  <div className={ styles.engageSearch }>
                    <input type="search" name="extnSearchBox" placeholder="Search for Engagements…" id={ styles.engageSearchBox }
                      onInputCapture={(evt) => this.setState({searchEngage: evt.target["value"].trim().toLowerCase()}) }
                    />
                    <i className="ms-Icon ms-Icon--Search" aria-hidden="true"></i>
                  </div>
                  <div className={` ${styles.extnContainer} ${styles.scrollHidden} ${styles.engageProgress2}`}>
                    {
                      
                    engagementsProcess1
                  }
                  </div>
                </div>
              </div>
              </div>
            </div>

            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} ${styles.colRecentDoc} `}>
              <div>
                <h4 className={styles.column2Title}>Shared Documents</h4>
                <div className={ styles.column2Container }>
                <div className={ styles.docTitle}>
                  <span>Name</span><span>Date Modified</span>
                </div>
                <div className={ styles.docContainer }>
                  // {myRecentDoc}
                  {mysharedDoc}
                </div>
                </div>
              </div>
            </div>
            
            <div className={`${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg8} ${styles.col4} ${styles.colProcessStage}`} style={{
              height: '400px', maxHeight: '400px'
            }}>
              <div>
                <h4 className={`${styles.column2Title}`}>PROCESS STAGES</h4>
                <div className={ styles.extnSearch }>
                  <input type="search" name="extnSearchBox" placeholder="Search for Process Stages…" id={ styles.extnSearchBox }
                    onInputCapture={({target}) => this.setState({searchProcessStage: target["value"].trim().toLowerCase()}) }
                  />
                  <i className="ms-Icon ms-Icon--Search" aria-hidden="true"></i>
                </div>
                {processesStage1}
              </div>
            </div>
            
            <div className={` ${styles.column2} ms-sm12 ms-lg8 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} ${styles.colSPSites} `} style={{
              height: '400px', maxHeight: '400px'
            }}>
              <div >
                <h4 className={`${styles.column2Title} ${styles.analyticsTitle}`}>Process Stage Analytics
                </h4>
                <div className={` ${styles.column2Container} ${styles.scrollHidden} `}>
                  <reactIframe.default
                    url={Data.biChartUrl[0]}
                    width="100%" height="800px"
                    styles={{display: "none"}}
                  />
                  
                </div>
              </div>
            </div>

            <div className={` ${styles.column2} ms-sm12 ms-lg8 ${styles.msSm12} ${styles.msLg8} ${styles.colTeams} ${styles.colSPSites} `} style={{
              height: '450px', maxHeight: '450px'
            }}>
              <div>
                <h4 className={`${styles.column2Title} ${styles.analyticsTitle}`}>MyANALYTICS
                  <button className={styles.moreButton}><img src={require('./images/ellipsis1.svg')} width="20px" /></button>
                </h4>
                <div className={` ${styles.column2Container} ${styles.scrollHidden} `} style={{overflowY: "scroll"}}>
                  
                  <div className={` ${styles.analyticsOverview} `}>
                    {Object.keys(this.state.CollaborationActivityList).map(activity => {
                      let duration = this.state.CollaborationActivityList[activity].reduce((acc, cur) => (acc + cur["parsedDuration"]), 0).toFixed(2), trackerColor = activity==="Meeting" ? "#e8c1a0" :  activity==="Focus" ? "#e8a838" :  activity==="Chats/Calls" ? "#f1e15b" :  activity==="Email" ? "#f47560" : "transparent";
                      
                      return (
                      <div className={` ms-sm12 ms-md3 ${styles.msSm12} ${styles.msMd3} ${styles.analyticsCard} `} >
                        <div>
                        <p style={{backgroundImage: `url('${require('./images/analyticsMeeting.svg')}')`}}>{activity}</p>
                        <p><span>{duration}</span><span>hrs</span><span className={`${styles.analyticsCardTracker}`}><span style={{width: `min(calc(100% * ${duration} / 24), 100%)`, backgroundColor: '#e8c1a0'}}></span></span></p>
                        </div>
                      </div>
                      );
                    })}
                  </div>

                  <div className={` ${styles.powerBiContainer} ${styles.analyticsContainer} `}>
                    <div className={` msSm12 msmd6  ${styles.msSm12} ${styles.msMd6}`} style={{display: "inline-block", position: "relative", left: 0}}>
                    {window.matchMedia("(min-width: 480px)").matches ? analyticsPieChart :  analyticsPieChartSM}
                    </div>
                    <div className={` msSm12 msmd6  ${styles.msSm12} ${styles.msMd6}`} style={{display: "inline-block", position: "relative", right: 0}}>
                    {analyticsBarChart1}
                    </div>
                  </div>
                </div>
              </div>
            </div>
            
            <div className={` ${styles.column2} ms-sm12 ms-lg8 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} ${styles.colSPSites} `} style={{
              height: '450px', maxHeight: '450px'
            }}>
              <div>
                <h4 className={`${styles.column2Title} ${styles.analyticsTitle}`}>Departmental Analytics
                  <button className={styles.moreButton}><img src={require('./images/ellipsis1.svg')} width="20px" /></button>
                </h4>
                <div className={` ${styles.column2Container} ${styles.scrollHidden} `}>
                  
                  <div className={` ${styles.powerBiContainer} ${styles.analyticsContainer} `} style={{height: 'calc(100% - 30px)'}}>
                    <div style={{width: 'calc(100% - 0px)', display: "inline-block", position: "relative", right: 0}}>
                    {analyticsBarChart2}
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colLauncher} `}>
              <div>
                <h4 className={`${styles.column2Title} ${styles.launcherTitle}`}>Launcher</h4>
                <div className={ styles.column2Container }>
                  <div className={` ms-Grid-row ${styles.msGridRow} ${styles.scrollHidden} `}>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--office ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/office_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/word" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--word ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/word_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/excel" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--excel ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/excel_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/powerpoint" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--powerpoint ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/powerpoint_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--access ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/access_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com/launch/onenote" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--onenote ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/onenote_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.com" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--onedrive ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/onedrive_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://www.microsoft.com/en-us/microsoft-365/project/project-management-software" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--project ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/project_96x1.png)'}}
                      ></div>
                    </a>
                    <a className={`ms-Grid-col ms-sm6 ms-md6 ms-lg4 ${styles.msGridCol} ${styles.msSm6} ${styles.msMd6} ${styles.msLg4}`} href="https://office.live.com/start/visio.aspx" target="_blank">
                      <div className={`ms-BrandIcon--icon96 ms-BrandIcon--visio ${styles.msBrandIcon96} ${styles.colLauncherIcon}`}
                        style={{backgroundImage: 'url(https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/png/visio_96x1.png)'}}
                      ></div>
                    </a>
                  </div>
                </div>
              </div>
            </div>


            

            */}
            




            {/* <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} ${ styles.colSPSites } `}>
              <div>
                <h4 className={`${styles.column2Title} ${styles.spTitle}`}>Employee Count</h4>
                <div className={` ${styles.column2Container} ${styles.scrollHidden} `}>
                  <div className={` ${styles.powerBiContainer} `}>
                    <ChartControl
                      type={ChartType.Bar}
                      options={barChartOptions}
                      data={barChartData}
                    />
                  </div>
                </div>
              </div>
            </div>

            <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colTwitter} `}>
              <div>
                <div className={ styles.twitterEmbed } dangerouslySetInnerHTML= {{__html:
                `<a class="twitter-timeline" data-width="100%" data-height="100%" data-tweet-limit=5 data-chrome="nofooter" href="https://twitter.com/RelianceInfoSys?ref_src=twsrc%5Etfw">Tweets</a>`
                }}>
                </div>
              </div>
            </div>

           <div className={`${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg8} ${styles.col4} ${styles.colProcessStage}`} style={{
              height: '400px', maxHeight: '400px'
            }}>
              <div>
              <div>
              <h4 className={`${styles.column2Title}`}>Employee TimeSheet</h4>
              <input id={ styles.extnSearchBox } type="text" placeholder="Search Name"></input>
             <button>Search</button>
              </div> 
               
               <table>
                {/*<th>Created</th>
               <th>User</th>
               <th>Period</th>
               <th>Period Starts</th>
               <th>Status</th><br></br>    */}
               {/* {this.state.items.map(function(item,key){
                  console.log(item);
                    return (<div  key={key}>
      
                      <tr> 
                        <td>{item.Created}</td>
                        <td>{item.Employee}</td>
                        <td>{formatDateTime(item.Period)}</td>
                        <td>{formatDateTime(item.Period_x002d_Starts)}</td>
                        <td>{item.TotalHours}</td>
                        <td>{item.Status}</td>
                        <td><FontAwesomeIcon icon={item.Status.toString()==="Approved"?'check-circle':item.Status.toString()==="Rejected"?'window-close':'check-circle'} style={{color:item.Status.toString()==="Approved"?'Green':item.Status.toString()==="Rejected"?'Red':'Green'}} ></FontAwesomeIcon></td>
                      </tr>    
                    </div>);
                    
                })}
             
            </table>      
              </div>
              </div>

              <div className={`${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg8} ${styles.col4} ${styles.colProcessStage}`} style={{
                height: '400px', maxHeight: '400px'
              }}>
                <div>
                  <h4 className={`${styles.column2Title}`}>TIME SHEET</h4>
                  <div className={ styles.extnSearch }>
                    <input type="search" name="extnSearchBox"placeholder="Search for Time Sheet…" id={ styles.extnSearchBox }
                      onInputCapture={({target}) => this.setState({searchTimeSheet: target["value"].trim().toLowerCase()}) }
                    />
                    <i className="ms-Icon ms-Icon--Search" aria-hidden="true"></i>
                  </div>
                  {timeSheet}
                </div>
              </div>

              <div className={` ${styles.column2} ms-sm12 ms-lg4 ${styles.msSm12} ${styles.msLg4} ${styles.colTeams} ${styles.colFAQ} `}>
                <div>
                  <h4 className={styles.column2Title}>FAQ</h4>
                  <div className={ styles.column2Container }>
                    <div className={ styles.faqHeaderImg }></div>
                    <div className={` ${styles.faqContent} ${styles.scrollHidden} `}>
                      {spFAQ}
                    </div>                  
                  </div>
                </div>
              </div>
              <div className={` ${styles.column2} ms-sm12 ms-lg8 ${styles.msSm12} ${styles.msLg12} ${styles.colTeams} ${styles.colSPSites} `} style={{
                height: '450px', maxHeight: '450px'
              }}>
                <div>
                  <h4 className={`${styles.column2Title}`}>Product Catalogue</h4>

                  <div className={`${styles.listSearch} ${styles.scrollHidden}`}>
                    <div className={styles.Pcontainer}>
                      <div className={styles.Prow}>
                        <div className={styles.Pcolumn}>
                          <span className={styles.title}>{this.props.list}</span>

                          <div className={styles.PmainContainer}>
                            <section className={styles.filterContainer}>
                              <h4>Filter</h4>
                              <hr />
                              <span className={styles.Ptitle}>
                                Products <br></br>{" "}
                              </span>
                              <hr />
                              {options.map((checkBoxItem: ICheckboxInput) => {
                                return (
                                  <Stack tokens={stackTokens}>
                                    <Checkbox
                                      className={styles.Plabel}
                                      label={checkBoxItem.Title}
                                      title={checkBoxItem.Title}
                                      onChange={(ev, checked) => this._onChange(ev, checked, "Product")}
                                    />
                                    <span></span>
                                  </Stack>
                                );
                              })}
                              <hr />
                              <span className={styles.Ptitle}>Existing Customers </span>
                              <hr />
                              {options1.map((checkBoxItem: ICheckboxInput) => {
                                return (
                                  <Stack tokens={stackTokens1}>
                                    <Checkbox
                                      label={checkBoxItem.Title}
                                      title={checkBoxItem.Title}
                                      onChange={(ev, checked) => this._onChange(ev, checked, "Customer")}
                                    />
                                    <span></span>
                                  </Stack>
                                );
                              })}
                              <hr />
                              <span className={styles.Ptitle}>
                                Sectors <br></br>{" "}
                              </span>
                              <hr />
                              {options2.map((checkBoxItem: ICheckboxInput) => {
                                return (
                                  <Stack tokens={stackTokens2}>
                                    <Checkbox
                                      label={checkBoxItem.Title}
                                      title={checkBoxItem.Title}
                                      onChange={(ev, checked) => this._onChange(ev, checked, "Sector")}
                                    />
                                    <span></span>
                                  </Stack>
                                );
                              })}
                              <hr />
                              <span className={styles.Ptitle}>
                                Industries <br></br>{" "}
                              </span>
                              <hr />
                              <label className={styles.Plabel}>  </label>
                              {options3.map((checkBoxItem: ICheckboxInput) => {
                                return (
                                  <Stack tokens={stackTokens3}>
                                    <Checkbox
                                      label={checkBoxItem.Title}
                                      title={checkBoxItem.Title}
                                      onChange={(ev, checked) => this._onChange(ev, checked, "Industry")}
                                    />
                                    <span></span>
                                  </Stack>
                                );
                              })}
                            </section>
                            <section className={styles.catalogContainer}>
                              <div className={`styles.searchContainer`}>
                                <div ref={this._menuButtonElement}>
                                  <SearchBox
                                    className={styles.PsearchBoxDiv}
                                    placeholder={`Search Product`}
                                    onSearch={(value) => this.searchListForContent(value)}
                                    onChanged={(value) => {
                                      this.setState({ PsearchText: value });
                                      this._onCalloutDismiss(value);
                                    }}
                                    onClear={(e) => this._onCalloutDismiss(e)}
                                  />
                                </div>
                              </div>{" "}
                              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                              <div className={styles.productContainerP}>
                                {CatalogCardArray}
                              </div>
                            </section>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            */}

          </div>
      );
    }
    else{
      return (
        <div></div>
      );

    }
  }
  
  
  public componentDidMount(): void{
    // console.log(`09/07/20 07:21AM`);
    // alert("Hi");
    this.checkConnectionState();
    if(this.props.context.sdks.microsoftTeams){
      this.setState({isOnTeams: true});
    }

    // this.carouselInterval = setInterval(() => {
    //   this.autoplayELements();
    // }, 3000);

      /* const script = document.createElement("script");
      script.src = "https://platform.twitter.com/widgets.js";
      document.getElementsByClassName(`${styles.twitterEmbed}`)[0].appendChild(script); */

      this._getCurrentUser();
      this._getMyTeamsGroup();


      this.RecurringFunction();
      
       
      //Carousel.call(this.);
      // var reactHandler = this;
  }

  public async RecurringFunction(){
    // this.checkConnectionState();
    this._getMyMailsMessage();
     this._getSPAnnouncement();
    this._getMyRecentDriveFiles();
    this.GetMyAnalytics();
    this.getLOBLinks();
    this._getOutlookEvents();
    this._getRecentEvent4rmList();
    this.getBirthdays();
    this.getAnniversary();

    this._getTasksAssignedToMe();

    window.setInterval(() => this.checkConnectionState(), 20000);

    window.setInterval(() => {
      this._getMyMailsMessage();
      //this._getMyTeamsGroup();
      this._getOutlookEvents();
      this._getRecentEvent4rmList();
      this.getBirthdays();
      this.getAnniversary();
      this._getTasksAssignedToMe();

      //console.log(staffAnniversaryThisMonth)

    }, 60000);


    window.setInterval(() => {
      this._getMyRecentDriveFiles();
    }, 240000);


  }

  // private triggerNextElement = (index: number) => {
  //   if (index > this.state.carouselItemElements.length-1) { // Beyond last element? Start with 1st one
  //     index = 0;
  //   }
  //   if (index < 0) { // Before 1st element? Go on with last one
  //     index = this.state.carouselItemElements.length-1;
  //   }
  //   this.setState({
  //     currentCarouselItemElement: this.state.carouselItemElements[index],
  //     currentCarouselItemIndex: index,
  //     carouselCanMoveNext: index < this.state.carouselItemElements.length-1,
  //     carouselCanMovePrev: index > 0
  //   });
  // }

  // private autoplayELements = () => {
  //   setTimeout(() => {
  //     this.triggerNextElement(this.state.currentCarouselItemIndex + 1);
  //   }, 1000);
  // }

  public checkConnectionState(){
    
    let xhr = new XMLHttpRequest();
    return new Promise((resolve, reject)=>{
      xhr.onload = () => this.setState({isConnected: true});
      xhr.onerror = () => this.setState({isConnected: false});
      xhr.open('GET', `${this.props.siteUrl}/_api/web/currentuser`, true);
      xhr.send();
    });
  }

  public _getCurrentUser(): void{
    var today = new Date(), later = new Date();
    later.setFullYear(today.getFullYear() + 1);
    
    this.myNotificationCount();
    
    this._msGraphFactory.then((clent: MSGraphClient): void => {
      clent.api('/me/').get((error: any, resp: IGraphUserProfile) =>{
        if(error){
          //console.log("GetCurrentUser Error", error);
          // if(this.state.isConnected){}
            this._getCurrentUser();
          return;
        }
        const result = new GraphUserProfile(resp);
        this.setState({myProfile: result});
      });
    });
    this._msGraphFactory.then((client: MSGraphClient):void => {
      client.api('/users?$top=999')
      .select(["id","mail","displayName","jobTitle","mail","mobilePhone","department","userPrincipalName","businessPhones"])
      // .filter()
      .get((error, resp: IGraphUserProfileItems) => {
        if(error){
          //console.log(error);      
          return;
        }
        const result = resp.value.map(user => {
          const newUser = new GraphUserProfile(user);
          /* newUser.dummyBirthday = this.utilityMethod.randomDate(today, later);
          newUser.manager = reportToArr[Math.floor(Math.random() * 3)]; */
          return newUser;
        });
        // console.log(`Employee Directory`, result);
        this.setState({allUser: result});
      });
    });
    this._msGraphFactory.then((client: MSGraphClient):void => {
      // client.api('/users').get((error, resp: IGraphUserProfileItems) => {
      client.api(`/me/people/?$filter=personType/class eq 'Person' and personType/subclass eq 'OrganizationUser'`)
      .version("v1.0").top(12)
      .get((error, resp: IGraphUserProfileItems) => {
        if(error){
          //console.log("UserError", error);
          return;
        }
        const result = resp.value.map(user => new GraphUserProfile(user));
        // console.log(result);
        this.setState({myRecentUsers: result});
      });
    });
    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/siteusers?$select=Id,Title,Email`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => response.json())
    .then((resp: {value: ISPUsers[]}) => {
      const result = resp.value.map(users => new SPUsers(users));
      
      this.setState({spUsers: result});
    })
    .catch(error => {
      if(this.state.isConnected){
        this._getCurrentUser();
      }
    });
  }

  public myNotificationCount(): void{
    var now = new Date(), later = moment.utc().add(7, "days");

    this._msGraphFactory.then((clent: MSGraphClient): void => {
      clent.api(`/me/reminderView(startDateTime='${now.toISOString()}',endDateTime='${later.toISOString()}')`).get((error: any, resp: {value: any[]}) =>{
        if(error){
          //console.log("GetNotifictionCount Error", error);
          if(this.state.isConnected){
            window.setTimeout(()=>this.myNotificationCount(), 120000);
          }
          return;
        }
        // console.log(`Notification COUNT`, resp);
        this.setState({notificationCount: resp.value.length});
      });
    });
  }

  public _getMyMailsMessage(): void{
    this._msGraphFactory.then((client: MSGraphClient): void =>{
      client.api('/me/mailFolders/Inbox/Messages')
      .select(["id", "receivedDateTime", "hasAttachments", "subject", "bodyPreview", "webLink", "from"])
      .filter('isRead ne true')
      .top(30)
      .count(true)
      .get((error: any, mailMessage: IGraphMailItems) => {
        if(error){
          //console.log("MailError", error);
          if(this.state.isConnected){
            this._getMyMailsMessage();
          }
          return;
        }
        // console.log("Messages", mailMessage);
        var mailArr = mailMessage.value.map(item => new GraphMail(item));

        this.setState({
          mailMessageArr: mailArr,
          mailMessageCount: mailMessage["@odata.count"]
        });
        
      });
    });
  }

  public async _getMyTeamsGroup(): Promise<any>{
    await this._msGraphFactory.then((client: MSGraphClient): void => {
      client.api('/me/joinedTeams').select(["id", "displayName"])
      .get( async (error: any, response: IGraphMyTeamItems) => {
        if(error){
          //console.log("TeamsError", error);
          // if(this.state.isConnected){}
            this._getMyTeamsGroup();
            
          return;
        }
        const resp = await response;
        // console.log(resp);
        const myJoinedTeams = await resp.value.map(async (item) => await this._getTopGroupIdInTeam(item));
      });
    });
  }

  // public _getGalleryImages(carouselItems: ICarouselImage[]): string[]{

  //   var images:string[] = [];
  //   //var resouceLinks: string [] = [];
  //   var i: number = 0;
  //   for(i = 0; i < carouselItems.length; i++){
  //      // iCarouselImage.push({FileRef:this._currentWebUrl.split("/sites")[0] + carouselItems[i].FileRef, ResourceLink: this./////_currentWebUrl.split("/sites")[0] + carouselItems[i].ResourceLink });
  //       images.push( carouselItems[i].FileRef);
  //       //resouceLinks.push(this._currentWebUrl.split("/sites")[0] + carouselItems[i].ResourceLink);
  //   }
  //   return images;
  // }

  // public _getGalleryResourseLink(carouselItems: ICarouselImage[]): string[]{
    
  //   var resouceLinks: string [] = [];
  //   var i: number = 0;
  //   for(i = 0; i < carouselItems.length; i++){
  //      // iCarouselImage.push({FileRef:this._currentWebUrl.split("/sites")[0] + carouselItems[i].FileRef, ResourceLink: this./////_currentWebUrl.split("/sites")[0] + carouselItems[i].ResourceLink });
  //      resouceLinks.push( carouselItems[i].ResourceLink);
  //       //resouceLinks.push(this._currentWebUrl.split("/sites")[0] + carouselItems[i].ResourceLink);
  //   }
  //   return resouceLinks;
  // }

  public _getTopGroupIdInTeam(team: IGraphMyTeam): void{
    this._msGraphFactory.then((client: MSGraphClient) => {
      let result: GraphMyTeam;
      client.api(`/teams/${team.id}/channels`)
      .get( async (error, resp: IGraphMyTeamItems) => {
        if(error){
          //console.log(team);
          //console.log("TeamError", error);
          // if(this.state.isConnected){}
            this._getTopGroupIdInTeam(team);
            
          return;
        }
        result = await new GraphMyTeam(team, resp.value[0].id, resp.value[0].webUrl);
        // console.log(result);
        if(this.state.myTeamGroupsArr.length === 0){
          this._getMyTeamMessages(result.teamId, result.groupId);
        }
        
        this.setState({myTeamGroupsArr: [...this.state.myTeamGroupsArr, result]});
      });
      
      return result;
    });
  }

  public _getMyTeamMessages(_teamId: string, _groupId: string):void {
    this._msGraphFactory.then((client: MSGraphClient) => {
      client.api(`/teams/${_teamId}/channels/${_groupId}/messages`)
      .version('beta')
      .top(50)
      .get(async (error: any, resp: IGraphTeamMessageItems) => {
        if(error){
          //console.log("TeamsError", error);
          if(this.state.isConnected){
            this._getMyTeamMessages(_teamId, _groupId);
          }
          return;
        }
        var a = await resp;
        
        // console.log(resp);
        // console.log(a);
        var teamMessage = resp.value.map(item => new GraphTeamMessage(item));
        // console.log(teamMessage);
        this.setState({selectedTeamMessages: teamMessage});
      });
    });
  }

  
  private _onChange(ev: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked: boolean, kind?: string) {
    // console.log(`The option ${ev.currentTarget.title} has been changed to ${isChecked}.`);
    let title = ev.currentTarget.title.toLowerCase();
    let productArr = (kind === "Product") ? this.state.productChecked : (kind === "Customer") ? this.state.customerChecked : (kind === "Sector") ? this.state.sectorChecked : (kind === "Industry") ? this.state.industryChecked : [];

    if(isChecked){
      productArr.push(title);
    }else{
      productArr.splice(productArr.indexOf(title), 1);
    }

    (kind === "Product") ? this.setState({productChecked: productArr}) : (kind === "Customer") ? this.setState({customerChecked: productArr}) : (kind === "Sector") ? this.setState({sectorChecked: productArr}) : (kind === "Industry") ? this.setState({industryChecked: productArr}) : console.log(null);
  }

  private _PsearchClicked(PsearchText: string): void {  
    //_getEnterpriseSearchWithFilterLimit();
   
    if(this.state.PsearchText.toString() == ""){
      this.setState({PsearchResult: [], Psearchstatus:false});      
    }else{
      this.searchListForContent(PsearchText);
      this.setState({PisCalloutVisible: true});
      //console.log(this.state.PisCalloutVisible);
    }
  
  }

  private searchListForContent(PsearchText: string){
    let queryText:string = escape(PsearchText);
    //console.log(queryText);

    const listQuery: string = `${this.props.siteUrl}/_api/web/lists/getByTitle('${this.props.listTitle}')/items?$select=*,Author/ID,Author/Title,Author/EMail,Author/Department,Editor/ID,Editor/Title,Editor/EMail,Editor/Department,AttachmentFiles/FileName,AttachmentFiles/ServerRelativeUrl&$expand=Author,Editor,AttachmentFiles&$filter=substringof('${queryText}', TargetMarkets) or substringof('${queryText}', Title) or substringof('${queryText}', ExistingCustomers)`;
    // let listQuery2:string = `${this.props.siteUrl}/_api/search/query?querytext='${searchText}*+Path:${this.props.siteUrl}/site/Lists/${this.props.listTitle}'`
    SPSearchService.get(this.props.context, listQuery)
    .then((resp) => {
      let result = resp.value.map(searchRes => new PSearchResult(searchRes) );
      
      //console.log(this.state.PisCalloutVisible);
      this.setState({PsearchResult:result, Psearchstatus:true, PisCalloutVisible: true});
      //console.log(this.state.PisCalloutVisible);
    });
  }
  
  private _onCalloutDismiss = (e): void => {
    this.setState({PisCalloutVisible: false});
    // if(e.target.className.indexOf("chartjs")<0){
      //console.log(`DISMISS started`, e.target);
      this.setState({
        PisCalloutVisible: false,
      });
    // }
  }

  private openResultDetails(reciordID: number){}

  public _getMyRecentDriveFiles(): any{
    
    this._msGraphFactory.then((client: MSGraphClient) => {
      client.api('/me/drive/recent').version("v1.0")
      .top(20)
      .get((error: any, resp: IGraphDriveFileItems) => {
        var result = resp;
        if(error){
          if(this.state.isConnected){
            window.setTimeout(()=>this._getMyRecentDriveFiles(), 120000);
          }
          return;
        }
        // console.log("MyRecentDoc", resp);

        var recentDoc = result.value.map(item => new GraphDriveFile(item));
        this.setState({myRecentDriveFiles: recentDoc});
      });
    });
  }

  public getLOBLinks(){

    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('LOB Apps')/items?$select=AttachmentFiles/ServerRelativeUrl,ApplicationLink,Title&$expand=AttachmentFiles`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => response.json())
    .then((resp: ILobLinkItems) => {
      // console.log(`LOB-Links RESPONSE`, resp.value);
      let result = resp.value.map(link => new LobLink(link));
      // console.log(`LOB-Links RESPONSE`, result);
      this.setState({LOBLinkList: result});
    })
    .catch(error => {
      //console.log("LOBLinks Error", error);
      if(this.state.isConnected){
        window.setTimeout(()=>this.getLOBLinks(), 120000);
      }
    });

  }

  public _getRecentEvent4rmList(){
    var today = new Date(), later = new Date();
    later.setDate(today.getDate() + 7);

    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Events')/items?$filter=EventDate ge datetime'${today.toISOString()}' or EndDate ge datetime'${today.toISOString()}'&$orderBy=EventDate asc&$top=15`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => response.json())
    .then((resp: ISPEventItems) => {
      //console.log(resp);
      const result = resp.value.map(event => new SPEvent(event));
      //console.log(result);
      this.setState({companyEvents: result});
    })
    .catch(error => {
      if(this.state.isConnected){
        this._getRecentEvent4rmList();
      }
    });

  }
  public utilityMethods = {
    monthOfTheYear: ["January","February","March","April","May","June","July","August","September","October","November","December"],
    mnthsOfTheYear: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul","Aug", "Sep", "Oct", "Nov", "Dec"],
  };

  public getBirthdays(): void {
    //var today = new Date(), later = new Date();
    //later.setDate(today.getDate() + 7);

    //this.props.spHttpClient.get(`${this.props.siteUrl2}/_api/web/lists/getbytitle('lst_Staff_Details')/items?$filter=Date_of_Birth ge datetime'${today.toISOString()}'&$orderBy=Date_of_Birth asc&$top=15`, SPHttpClient.configurations.v1)
    this.props.spHttpClient.get(`${this.props.siteUrl2}/_api/web/lists/getbytitle('lst_Staff_Details')/items?$Top=500&$orderby=Date_of_Birth asc`, SPHttpClient.configurations.v1)
    .then(resp => resp.json())
    .then((response: ISPEventItem) => {
      const result = response.value.map(listItems => new SPEvents(listItems));

      //console.log(result);
      //console.log(`staffBirthdays`, result);
      this.setState({staffBirthdays: result});
    });
  }
    //.then((response: SPHttpClientResponse) => response.json())
    //.then((resp: any) => {
     //const result = resp.value.map(bday => new SPEvent(bday));
      //console.log("staffBirthdays: ", resp);
      //this.setState({staffBirthdays: resp.value});
   // })
    //.catch(error => {
      //if(this.state.isConnected){
       // this.getBirthdays();
     // }
    //});

  //}

  public getAnniversary(): void {
    //var today = new Date(), later = new Date();
    //later.setDate(today.getDate() + 7);

   // this.props.spHttpClient.get(`${this.props.siteUrl2}/_api/web/lists/getbytitle('lst_Staff_Details')/items?$filter=Resumption_Date ge datetime'${today.toISOString()}'&$orderBy=Resumption_Date asc&$top=15`, SPHttpClient.configurations.v1)
   this.props.spHttpClient.get(`${this.props.siteUrl2}/_api/web/lists/getbytitle('lst_Staff_Details')/items?$Top=409&$orderby=Resumption_Date asc&$select=*,Staff_Name/EMail&$expand=Staff_Name`, SPHttpClient.configurations.v1)
   .then(resp => resp.json())
    .then((response: ISPEventItem) => {
      const result = response.value.map(listItems => new SPEvents(listItems));

      // console.log(`staffAnniversary`, result);
      this.setState({staffAnniversary: result});
      
    console.log('staffAnniversaryResult', result, this.state.staffAnniversary)
    });
    console.log('staffAnniversary', this.state.staffAnniversary)
  }
    //.then((response: SPHttpClientResponse) => response.json())
    //.then((resp: any) => {
      // const result = resp.value.map(bday => new SPEvent(bday));
      //console.log("staffAnniversary: ", resp);
     // this.setState({staffAnniversary: resp.value});
    //})
   // .catch(error => {
     // if(this.state.isConnected){
     //   this.getBirthdays();
     // }
  //  });

  //}

  public _getOutlookEvents(): any{
    var today = new Date(), later = new Date();
    later.setDate(today.getDate() + 7);
    // subject,body,bodyPreview,organizer,attendees,start,end,location
    const url = `/me/calendarview?startdatetime=${today.toISOString()}&enddatetime=${later.toISOString()}`;
    // console.log(url);

    this._msGraphFactory.then(client => {
      client.api(url)
      .select(["subject","body","bodyPreview","organizer","attendees","start","end","location","webLink", "isOnlineMeeting", "onlineMeeting"])
      .get((err, result: {value: MicrosoftGraph.Event[]}) => {
        // console.log("OutlookCalendar", result);
        // if(result != undefined){
          // console.log("OutlookCalendar", result);
          // console.log("OutlookCalendarError", err);
          const events = result.value.map(res => new myOutlookEvent(res));
          this.setState({myCalendar: events});
        // }
      });
    });
  }

  public _getTasksAssignedToMe(): any{
    
    this._msGraphFactory.then(client => {
      client.api(`/me/planner/tasks`)
      .get((error, tasks: {value: MicrosoftGraph.PlannerTask[]}) => {

        //console.log(`TASKS`, this.state.myProfile.mail, tasks);
        this.setState({myPlannerTasks: tasks.value});
      });
    })
    .catch(error => {
      if(this.state.isConnected){
        this._getTasksAssignedToMe();
      }
    });
  }

  private _PonCalloutDismiss = (e): void => {
    if(e.target.className.indexOf("chartjs")<0){
      //console.log(`DISMISS started`, e.target);
      this.setState({
        isCalloutVisible: false,
      });
    }
  }
  public updateTaskPercent(taskId: string, taskOdata: string, taskPercent: number = 100){
    //console.log(taskId, taskOdata, taskPercent);
    this._msGraphFactory.then(client =>{
      client.api(`planner/tasks/${taskId}`)
      .headers({'Content-type': 'application/json', 'If-Match': taskOdata })
      .patch({percentComplete: taskPercent})
      //.then(res => console.log(res))
     // .catch(err => console.log(err));
    })
  }

  private _searchClicked(): void {  
    //_getEnterpriseSearchWithFilterLimit();
   
    if(this.state.searchText.toString() == ""){
      this.setState({item:[],searchstatus:false});      
    }else{
      this._getEnterpriseSearchWithFilterLimit(this.state.searchText.toString(),10);
      this.setState({isCalloutVisible: true});
      //console.log(this.state.isCalloutVisible);
    }
  
   // debugger;
  }

  private _searchClearClicked(): void {  
    //_getEnterpriseSearchWithFilterLimit();
   
    if(this.state.searchText.toString() == ""){
      this.setState({item:[],searchstatus:false});      
    }else{
      this.setState({searchText:"",searchstatus:false});
    }
  
   // debugger;
  }

  public _getEnterpriseSearchWithFilterLimit(searchphrase: string ,limit : number){
    var baseUrl = this._getBaseUrl(this.props.siteUrl,this.props.subSitePath);
    //console.log(baseUrl);
    //debugger;
    //${baseUrl}&rowlimit=${limit.toString()}
    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/search/query?querytext='${searchphrase.toString()}'&amp;rowlimit=${limit.toString()}`, SPHttpClient.configurations.v1,
    { headers: {'odata-version': '3.0'}})
    .then((response: SPHttpClientResponse) => response.json())
    .then((resp) => {
      //console.log(resp);
      //const result = resp.value.map(event => new SPEvent(event));      
      //console.log(result);
      let searchResp: ISPSearchResult[] = [];  

      // Check if there was an error  
      if (typeof resp["odata.error"] !== "undefined") {  
        if (typeof resp["odata.error"]["message"] !== "undefined") {  
          Promise.reject(resp["odata.error"]["message"].value);  
          return;  
        }  
      }  

      if (!this._isNull(resp)) {  
        const fields: string = "Title,Path,Description,LastModifiedTime,FileExtension,Author";

        // Retrieve all the table rows  
        if (typeof resp.PrimaryQueryResult.RelevantResults.Table !== 'undefined') {  
          if (typeof resp.PrimaryQueryResult.RelevantResults.Table.Rows !== 'undefined') {                
            searchResp = this._setSearchResults(resp.PrimaryQueryResult.RelevantResults.Table.Rows, fields); 
            //console.log(searchResp);
          }  
        }  
      }  

      //console.log(this.state.isCalloutVisible);
      this.setState({item:searchResp,searchstatus:true, isCalloutVisible: true});
      //console.log(this.state.isCalloutVisible);
    })
    .catch(error => {
      //console.log(error);
      // if(this.state.isConnected){}
      //this._getRecentEvent4rmList();
    });

  }

  private _setSearchResults(crntResults: ICells[] , fields: string): any[] {  
    const temp: any[] = [];  

    if (crntResults.length > 0) {  
      const flds: string[] = fields.toLowerCase().split(',');  

      crntResults.forEach((result) => {  
        // Create a temp value  
        var val: Object = {} ;

        result.Cells.forEach((cell: ICellValue) => {  
          if (flds.indexOf(cell.Key.toLowerCase()) !== -1) {  
            // Add key and value to temp value  
            val[cell.Key] = cell.Value;  
          }  
        });  

        // Push this to the temp array  
        temp.push(val);
      });  
    }  

    return temp;  
  }

  private _isNull(value: any): boolean {  
    return value === null || typeof value === "undefined";  
  }
  private _getBaseUrl(siteUrl: string, subSitePath: string): string {
    var baseUrl = null;
    // console.log(subSitePath);
    baseUrl = siteUrl.replace(subSitePath, '');
    //baseUrl = baseUrl+"";
    return baseUrl;
  }


  public _getSPAnnouncement(){
    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Announcement')/items?$select=Id,Title,Description,AuthorId,AttachmentFiles/ServerRelativeUrl,Created,Author/Title&$orderby=Created desc&$expand=AttachmentFiles,Author`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => response.json())
    .then((resp: ISPAnnouncementItems) => {
      const result = resp.value.map(extn => new SPAnnouncement(extn));
      
      this.setState({spAnnouncements: result});
    })
    .catch(error => {
      if(this.state.isConnected){
        this._getSPAnnouncement();
      }
    });
  }

  public async getMySPId(): Promise<string>{
    const currentuser = await this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/currentuser?$select=Id`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => response.json());
    //console.log("USER!!!", currentuser);
    return currentuser["Id"];
  }

  
  public GetMyAnalytics(){
    this._msGraphFactory.then(client => {
      client.api('/me/analytics/activitystatistics')
      .version('beta').get((error, resp) => {
        //console.log("MyANALYTICS", resp.value);

        var Activities = new Array<Activity>(), groupActivities = {},colActy={}, aggActivities = [];

        Activities = resp.value;
        Activities.forEach((acty: Activity)=>{
          acty.parsedDuration = (toSeconds(parse(acty.duration))) / 3600;

          var testAcc = groupActivities[acty.activity];
          testAcc ? testAcc.push(acty) : groupActivities[acty.activity] = [acty];
        });

        Activities.forEach((acty: Activity)=>{
          if(acty.activity === "Call" || acty.activity === "Chat" ){
            var callandchatactivity = "Chat & Call";
            acty.parsedDuration = (toSeconds(parse(acty.duration))) / 3600;
            let testAcc = colActy[callandchatactivity];
            testAcc ? testAcc.push(acty) : colActy[callandchatactivity] = [acty];

          }else{
            acty.parsedDuration = (toSeconds(parse(acty.duration))) / 3600;

            let testAcc = colActy[acty.activity];
            testAcc ? testAcc.push(acty) : colActy[acty.activity] = [acty];
          }
        
        });
      

        Object.keys(groupActivities).forEach(activity => {
          let aggActy: AggregatedActivity = {
            activity: activity,
            activityList: groupActivities[activity],
            aggregatedDuration: 0, aggregatedAfterHours: 0
          };
          aggActy.aggregatedDuration = aggActy.activityList.reduce((acc, cur, i) => (acc + cur.parsedDuration), 0);
          aggActivities.push(aggActy);
        });
        
      this.setState({
        CollaborationActivityList:colActy,
        GroupActivityList: groupActivities,
        AggregatedActivityList: aggActivities,
        barChartData: this.formatBarChartData(groupActivities),
        // pieChartData: this.formatPieChartData(groupActivities)
        pieChartData: this.formatPieChartData2(groupActivities)
      });
      });
    });
  }
  
  public formatBarChartData(fullData: GroupedActivity){
    const dataKey = Object.keys(fullData);
    let result = {};

    dataKey.forEach(acty=> {
      fullData[acty].forEach(dayActy =>{
        if(result[dayActy["startDate"]]){
          result[dayActy["startDate"]][acty] =dayActy["parsedDuration"];
        } else{
          result[dayActy["startDate"]] = {DAY: this.utilityMethod.daysOfWeek[(new Date(dayActy["startDate"])).getDay()]};
          result[dayActy["startDate"]][acty]=dayActy["parsedDuration"];
        }
      });
    });

    // const fullBarChartData = Object.keys(result).map(key => result[key]).map(day => {
    const fullBarChartData = Object.keys(result).map(key => {
      let day = result[key], collabVal = 0;
      Object.keys(day).forEach(dayKey => collabVal = (dayKey==="Focus" || isNaN(day[dayKey])) ? collabVal + 0 : collabVal + day[dayKey] );
      return {...day, Collab: collabVal};
    });
    //console.log( fullBarChartData );
    return fullBarChartData;
  }
  
  public formatPieChartData2(fullData: GroupedActivity){
    // const Data = this.state.CollaborationActivityList;
    const Data1 = fullData;
    let activityKey = Object.keys(Data1), pieData: PieGraphData[] = [];

    pieData = activityKey.map(activity => {
      let duration: number = Data1[activity].reduce((acc, cur) => (acc + cur["parsedDuration"]), 0).toFixed(2);
      return {id: activity, value: duration};
    });

    return pieData;
  }

  public _teamSelectedListener(event): any {
    const et = event.target.attributes;
    // console.dir(et);
    this.setState({selectedTeam: et.id.value});
    this._getMyTeamMessages(et["data-teamid"].value, et["data-groupid"].value);
  }

  public _outlookMailListener(mailId: string){
    let newMailArr = this.state.mailMessageArr.filter(mail => mail.id !== mailId), newMailCount = this.state.mailMessageCount - 1;
    this.setState({mailMessageArr: newMailArr, mailMessageCount: newMailCount});
    
    this._msGraphFactory.then(client => {
      client.api(`/me/messages/${mailId}`)
      .update({isRead: true})
      .then(result => {
        //console.log(`Mark email as READ`, result);
        this._getMyMailsMessage();
      });
    });
  }
  public _outlookHeaderListener(evnt: Event){
    //console.log(evnt.type);
  }

  public utilityMethod = {
    daysOfTheWeek: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
    daysOfWeek : ["Sun", "Mon", "Tues", "Wed", "Thurs", "Fri", "Sat"],
    monthsOfTheYear: ["January","February","March","April","May","June","July","August","September","October","November","December"],
    mnthsOfTheYear: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul","Aug", "Sep", "Oct", "Nov", "Dec"],
    timeSince : (date: Date): string => {
      var seconds = Math.floor(((new Date()).getTime() - date.getTime()) / 1000);
      var interval = Math.floor(seconds / 31536000);
      var daysInMnth = [31,28,31,30,31,30,31,31,30,31,30,31];

      if (interval > 1) {
        return interval + " years ago";
      }
      interval = Math.floor(seconds / (60*60*24*daysInMnth[date.getMonth()]));
      if (interval > 1) {
        return interval + " months ago";
      }
      interval = Math.floor(seconds / (60*60*24));
      if (interval > 1) {
        return interval + " days ago";
      }
      interval = Math.floor(seconds / (60*60));
      if (interval > 1) {
        return interval + " hours ago";
      }
      interval = Math.floor(seconds / 60);
      if (interval > 1) {
        return interval + " minutes ago";
      }
      return Math.floor(seconds) + " seconds ago";
    },
    convertDateTime: (date: Date, separate: string, fullYear?: Boolean, monthInString?: Boolean, monthInFull?: Boolean, showYear?: Boolean): string => `${date.getDate() < 10 ? "0" : ""}${date.getDate()}${separate}${(date.getMonth()+1 < 10 && !monthInString) ? "0" : ""}${!monthInString ? date.getMonth()+1 : monthInFull ? this.utilityMethod.monthsOfTheYear[date.getMonth()] : this.utilityMethod.mnthsOfTheYear[date.getMonth()]}${!showYear ? "" :separate}${!showYear ? "" :fullYear ? date.getFullYear() : date.getFullYear().toString().substr(-2)}`,
    getFormattedTime: (dateTime: Date): string => `${dateTime.getHours() === 0 ? `12` : dateTime.getHours() > 12 ? dateTime.getHours() - 12 : dateTime.getHours()}:${dateTime.getMinutes() < 10 ? "0" : ""}${dateTime.getMinutes()}${dateTime.getHours() === 0 ? `AM` : dateTime.getHours() > 12 ? 'PM' : 'AM'}`,
    getCurrentSPUser: async(): Promise<any> => {
      const currentUser = await this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/currentuser`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
      return currentUser;
    },
    randomDate: (start: Date, end: Date, startHour?: number, endHour?: number) => {
      var date = new Date(Number(start) + Math.random() *(Number(end) - Number(start)));
      var hour = startHour + Math.random()*(endHour - startHour) | 0;
      date.setHours(hour);
      return date;
    }
  };
}
