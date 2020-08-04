import * as React from 'react';
import { ICustomPeopleDirectoryProps } from './ICustomPeopleDirectoryProps';
import { ICustomPeopleDirectoryState } from './ICustomPeopleDirectoryState';
import { MSGraphClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { appendUserPhoto } from './photoFunction';
import { Modal, IconButton, Stack, Image, IImageProps, ImageFit, Pivot, PivotItem } from 'office-ui-fabric-react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import Masonry from 'react-masonry-css';

require('./style.css');
let AllUserList: any = [];
let searchUserList: any = [];
let colleagueDetails: any = [];
let managerDetails: any = [];

export default class CustomPeopleDirectory extends React.Component<ICustomPeopleDirectoryProps, ICustomPeopleDirectoryState> {

  public constructor(props: ICustomPeopleDirectoryProps, state: ICustomPeopleDirectoryState) {
    super(props);

    // STATE INITIALIZATION

    this.state = {
      UserListData: [],
      HideShowSearchIcon: false,
      HideModalPop: false,
      CurrentUserId: "",
      CurrentUserDisplayName: "",
      CurrentUserTitle: "",
      CurrentUserDOJ: "",
      CurrentUserEmail: "",
      CurrentPhotoLink: "",
      CurrentWorkPhone: "",
      CurrentMobilePhone: "",
      CurrentOfficeLocation: "",
      CurrentWorkDepartment: "",
      CurrentManagerDisplayName: "",
      CurrentManagerDOJ: "",
      CurrentManagerPhotoLink : "",
      CurrentColleagueName: [],
      SearchTextValue: ""
    };

  }

  public _context: WebPartContext;

  public componentDidMount() { this.InvokeUserFetch(); }

  async InvokeUserFetch() {

    /* await this.props.graphClient.getClient().then((client: MSGraphClient): void => {
      client.api('/users').top(999).get().then((res) => {
        for (let i = 0; i < res.value.length; i++) {
            let blobUrl = await appendUserPhoto( client,res.value[i].id );
            let userData = {
              "id": res.value[i].id,
              "userPrincipalName": res.value[i].userPrincipalName,
              "DisplayName": res.value[i].displayName,
              "DOJ": res.value[i].jobTitle,
              "email": res.value[i].mail,
              "photoLink": blobUrl
            }
            userList.push(userData);
        }
        this.setState({ UserListData: userList });
      }).catch((err) => {
        console.log(err);
      });
    }).catch((err) => {
      console.log(err);
    }); */

    let client: MSGraphClient = await this.props.graphClient.getClient();
    let _filterQuery = "(accountEnabled eq true)";
    let cols = ["businessPhones", "displayName", "givenName", "jobTitle", "mail", "mobilePhone", "officeLocation", "id", "department"]
    let userList = await client.api('/users').top(999).filter(_filterQuery).select(cols).get();

    for (let i = 0; i < userList.value.length; i++) {
      if (userList.value[i].department != null) {
        let userId = ""; let userPrincipalName = ""; let displayName = "";
        let jobTitle = ""; let mail = ""; let businessPhones = ""; let mobilePhone = ""; let officeLocation = ""; let department = "";

        // let blobUrl = await appendUserPhoto(client, userList.value[i].userPrincipalName);
        if (userList.value[i].id != null) {
          userId = userList.value[i].id;
        }

        if (userList.value[i].department != null) {
          department = userList.value[i].department;
        }

        if (userList.value[i].userPrincipalName != null) {
          userPrincipalName = userList.value[i].userPrincipalName;
        }

        if (userList.value[i].displayName != null) {
          displayName = userList.value[i].displayName;
        }

        if (userList.value[i].jobTitle != null) {
          jobTitle = userList.value[i].jobTitle;
        }

        if (userList.value[i].mail != null) {
          mail = userList.value[i].mail;
        }

        if (userList.value[i].businessPhones[0] != null) {
          businessPhones = userList.value[i].businessPhones[0];
        }

        if (userList.value[i].mobilePhone != null) {
          mobilePhone = userList.value[i].mobilePhone;
        }

        if (userList.value[i].officeLocation != null) {
          officeLocation = userList.value[i].officeLocation;
        }

        let userData = {
          "id": userId,
          "userPrincipalName": userPrincipalName,
          "DisplayName": displayName,
          "DOJ": jobTitle,
          "email": mail,
          "photoLink": "/_layouts/15/userphoto.aspx?size=L&username=" + mail + "",
          "workNumber": businessPhones,
          "OfficeLocation": officeLocation,
          "mobileNumber": mobilePhone,
          "department": department
        };
        AllUserList.push(userData);
        AllUserList.sort((a, b) => (a.DisplayName > b.DisplayName) ? 1 : -1)
      }
    }
    this.setState({ UserListData: AllUserList });
  }

  async captureTextChange(enteredText) {
    searchUserList = [];
    if (enteredText != "") {
      let _filterQuery = "(startswith( DisplayName,'" + enteredText + "')) and (accountEnabled eq true)";
      let cols = ["businessPhones", "displayName", "givenName", "jobTitle", "mail", "mobilePhone", "officeLocation", "id", "department"]
      let client: MSGraphClient = await this.props.graphClient.getClient();
      let userList = await client.api('/users').top(999).filter(_filterQuery).select(cols).get();
      if (userList.value.length != 0) {
        for (let i = 0; i < userList.value.length; i++) {
          if (userList.value[i].department != null) {
            let userId = ""; let userPrincipalName = ""; let displayName = "";
            let jobTitle = ""; let mail = ""; let businessPhones = ""; let mobilePhone = ""; let officeLocation = ""; let department = "";

            // let blobUrl = await appendUserPhoto(client, userList.value[i].userPrincipalName);
            if (userList.value[i].id != null) {
              userId = userList.value[i].id;
            }

            if (userList.value[i].department != null) {
              department = userList.value[i].department;
            }

            if (userList.value[i].userPrincipalName != null) {
              userPrincipalName = userList.value[i].userPrincipalName;
            }

            if (userList.value[i].displayName != null) {
              displayName = userList.value[i].displayName;
            }

            if (userList.value[i].jobTitle != null) {
              jobTitle = userList.value[i].jobTitle;
            }

            if (userList.value[i].mail != null) {
              mail = userList.value[i].mail;
            }

            if (userList.value[i].businessPhones[0] != null) {
              businessPhones = userList.value[i].businessPhones[0];
            }

            if (userList.value[i].mobilePhone != null) {
              mobilePhone = userList.value[i].mobilePhone;
            }

            if (userList.value[i].officeLocation != null) {
              officeLocation = userList.value[i].officeLocation;
            }

            let userData = {
              "id": userId,
              "userPrincipalName": userPrincipalName,
              "DisplayName": displayName,
              "DOJ": jobTitle,
              "email": mail,
              "photoLink": "/_layouts/15/userphoto.aspx?size=L&username=" + mail + "",
              "workNumber": businessPhones,
              "OfficeLocation": officeLocation,
              "mobileNumber": mobilePhone,
              "department": department
            };
            searchUserList.push(userData);
          }
        }
        this.setState({ UserListData: searchUserList });
      }
    } else {
      this.setState({ UserListData: AllUserList });
    }
  }

  public render(): React.ReactElement<ICustomPeopleDirectoryProps> {
    let name: any;
    let letter: any;
    let colors: any;
    let colorSelector: any;

    if (this.state.CurrentManagerDisplayName != "") {
      name = this.state.CurrentManagerDisplayName.match(/\b(\w)/g);
      letter = name.join().replace(/,/g, '');
      colors = ["", "#ca5010", "#038387", "#005b70", "#986f0b", "#881798", "881798"];
      colorSelector = colors[Math.floor(Math.random() * (5 - 1)) + 1];
    }

    const imageProps: IImageProps = {
      imageFit: ImageFit.contain,
    };
    if (this.state.UserListData != 0) {
      return (
        <div>
          <section className="user_container">
            <section className="user_search_container">
              <div>
                <input placeholder="Who are you looking for ?" onChange={e => this.captureTextChange(e.target.value)} />
                <span className="user_search_icon_wrap">
                  <span>
                  </span>
                </span>
              </div>
            </section>
            <section className="user_list_container">
              <div className="user_card_container grid">
                {this.bindUserData(this.state.UserListData)}
              </div>
            </section>
          </section>
          <Modal isOpen={this.state.HideModalPop} onDismiss={this.closeModal.bind(this)} isBlocking={false}  >
            <Stack className="modal_popup" styles={{ root: { height: 600, width: 800 } }}>
              <div>
                <h2 className="text-right m-0"><IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel="Close popup modal" onClick={this.closeModal.bind(this)} style={{ marginRight: 10 }} /> </h2>
              </div>
              <Stack>
                <div className="d-flex align-items-center modal_header">
                  <div>
                    <Image src={this.state.CurrentPhotoLink} {...imageProps} width={100} height={100} className="b-r-r"></Image>
                  </div>
                  <div className="modal_user">
                    <p className="modal_user_name">{this.state.CurrentUserDisplayName}</p>
                    <p className="modal_user_date">{this.state.CurrentUserDOJ} </p>
                    <p className="modal_user_date"> {this.state.CurrentWorkDepartment}</p>
                  </div>
                </div>
                <div className="modal_body">
                  <Pivot>
                    <PivotItem
                      headerText="Overview">
                      <div className="modal_content">
                        <div className="modal_content_title">
                          <h4>Contact Information</h4>
                        </div>
                        <div className="modal_contact_info">
                          <div className="modal_info_card">
                            <div className="modal_info_icon">
                              <Icon iconName="Mail" />
                            </div>
                            <div className="modal_info_content">
                              <p>Email</p>
                              <p className="t-blue">{this.state.CurrentUserEmail}</p>
                            </div>
                          </div>
                          <div className="modal_info_card">
                            <div className="modal_info_icon">
                              <Icon iconName="Chat" />
                            </div>
                            <div className="modal_info_content">
                              <p>Chat</p>
                              <p className="t-blue">{this.state.CurrentUserEmail}</p>
                            </div>
                          </div>
                          <div className="modal_info_card">
                            <div className="modal_info_icon">
                              <Icon iconName="Phone" />
                            </div>
                            <div className="modal_info_content">
                              <p>Work Phone</p>
                              <p className="t-blue">{this.state.CurrentWorkPhone}</p>
                            </div>
                          </div>
                        </div>
                        <div className="modal_contact_info">
                          <div className="modal_info_card">
                            <div className="modal_info_icon">
                              <Icon iconName="Contact" />
                            </div>
                            <div className="modal_info_content">
                              <p>Job title</p>
                              <p className="t-blue">{this.state.CurrentUserDOJ}</p>
                            </div>
                          </div>
                          <div className="modal_info_card">
                            <div className="modal_info_icon">
                              <Icon iconName="Group" />
                            </div>
                            <div className="modal_info_content">
                              <p>Department</p>
                              <p className="t-blue">{this.state.CurrentWorkDepartment}</p>
                            </div>
                          </div>
                        </div>
                        <br></br>
                        <div className="modal_content_title">
                          <h4>Organization</h4>
                        </div>
                        <div className="org_wrap d-flex">
                          <div className="org_wrap_manager">
                            <span className="modal_you_work">Manager</span>
                            <div className="modal_org_content_wrap">

                              <div className="modal_org_card">
                                <div className="org_letter" >
                                  {/* {letter} */}
                                  <Image src={this.state.CurrentManagerPhotoLink} width={50} height={50} className="b-r-r"></Image>
                                </div>
                                <div className="modal_org_detail">
                                  <p>{this.state.CurrentManagerDisplayName}</p>
                                  <p>{this.state.CurrentManagerDOJ}</p>
                                </div>
                              </div>

                            </div>
                          </div>
                          <div className="org_wrap_users">
                            <span className="modal_you_work">You work with</span>
                            <div className="modal_org_content_wrap">
                              {this.state.CurrentColleagueName.map(function (data) {
                                const name = data.displayName.match(/\b(\w)/g),
                                  letter = name.join().replace(/,/g, ''),
                                  colors = ["", "#ca5010", "#038387", "#005b70", "#986f0b", "#881798", "881798"],
                                  colorSelector = colors[Math.floor(Math.random() * (5 - 1)) + 1];

                                return (
                                  <div className="modal_org_card">
                                    <div className="d-flex align-items-center">
                                      {/* <div className="org_letter" style={{ backgroundColor: colorSelector }}>{letter}</div> */}
                                      <div className="org_letter" >
                                        <Image src={data.photoLink} width={40} height={40} className="b-r-r"></Image>
                                      </div>
                                      <div className="modal_org_detail">
                                        <p>{data.displayName}</p>
                                        <p>{data.DOJ}</p>
                                      </div> 
                                    </div>
                                  </div>
                                )
                              })}
                            </div>
                          </div>
                        </div>
                        {/* <div className="modal_org_content_wrap">
                          <div className="modal_org_content">
                            <div className="modal_org_card">
                              <div className="modal_org_letter">
                                Manager
                              </div>
                              <div className="modal_org_detail">
                                <p>{this.state.CurrentManagerDisplayName}</p>
                                <p>{this.state.CurrentManagerDOJ}</p>
                              </div>
                            </div>
                          </div>
                          <div className="modal_org_content">
                            <div className="modal_org_card">
                              <div className="modal_org_letter">
                                Colleagues
                              </div>
                              {this.state.CurrentColleagueName.map(function (data) {
                                return (
                                  <div className="modal_org_detail">
                                    <p>{data.displayName}</p>
                                    <p>{data.DOJ}</p>
                                  </div>
                                );
                              })};
                            </div>
                          </div>
                        </div> */}
                      </div>
                    </PivotItem>
                    {/* <PivotItem headerText="Contact">
                      <div className="modal_content">
                        <div className="modal_content_title">
                          <h4>Contact Information</h4>
                        </div>
                        <div className="modal_contact_info">
                          <div className="modal_info_card">
                            <div className="modal_info_icon">
                              <Icon iconName="Mail" />
                            </div>
                            <div className="modal_info_content">
                              <p>Email</p>
                              <p className="t-blue">{this.state.CurrentUserEmail}</p>
                            </div>
                          </div>
                          <div className="modal_info_card">
                            <div className="modal_info_icon">
                              <Icon iconName="Chat" />
                            </div>
                            <div className="modal_info_content">
                              <p>Chat</p>
                              <p className="t-blue">{this.state.CurrentUserEmail}</p>
                            </div>
                          </div>
                          <div className="modal_info_card">
                            <div className="modal_info_icon">
                              <Icon iconName="Phone" />
                            </div>
                            <div className="modal_info_content">
                              <p>Work Phone</p>
                              <p className="t-blue">{this.state.CurrentWorkPhone}</p>
                            </div>
                          </div>
                        </div>
                        <div className="modal_contact_info">
                          <div className="modal_info_card">
                            <div className="modal_info_icon">
                              <Icon iconName="Contact" />
                            </div>
                            <div className="modal_info_content">
                              <p>Job title</p>
                              <p>{this.state.CurrentUserDOJ}</p>
                            </div>
                          </div>
                          <div className="modal_info_card">
                            <div className="modal_info_icon">
                              <Icon iconName="Group" />
                            </div>
                            <div className="modal_info_content">
                              <p>Department</p>
                              <p>{this.state.CurrentWorkDepartment}</p>
                            </div>
                          </div>
                        </div>
                      </div>
                    </PivotItem>
                    <PivotItem headerText="Organization">
                      <div className="modal_content">
                        <p>Organization Details</p>
                      </div>
                      <div className="org_wrap d-flex">
                        <div className="org_wrap_manager">
                          <span className="modal_you_work">Manager</span>
                          <div className="modal_org_content_wrap">
                            <div className="modal_org_card">
                              <div className="org_letter">AS</div>
                              <div className="modal_org_detail">
                                <p>{this.state.CurrentManagerDisplayName}</p>
                                <p>{this.state.CurrentManagerDOJ}</p>
                              </div>
                            </div>
                          </div>
                        </div>
                        <div className="org_wrap_users">
                          <span className="modal_you_work">You work with</span>
                          <div className="modal_org_content_wrap">
                            {this.state.CurrentColleagueName.map(function (data) {
                              const name = data.displayName.match(/\b(\w)/g),
                                letter = name.join().replace(/,/g, ''),
                                colors = ["", "#ca5010", "#038387", "#005b70", "#986f0b", "#881798", "881798"],
                                colorSelector = colors[Math.floor(Math.random() * (5 - 1)) + 1];
                              return (
                                <div className="modal_org_card">
                                  <div className="d-flex align-items-center">
                                    <div className="org_letter" style={{ backgroundColor: colorSelector }}>{letter}</div>
                                    <div className="modal_org_detail">
                                      <p>{data.displayName}</p>
                                      <p>{data.DOJ}</p>
                                    </div>
                                  </div>
                                </div>
                              )
                            })}
                          </div>
                        </div>
                      </div>
                    </PivotItem> */}
                  </Pivot>
                </div>
              </Stack>
            </Stack>
          </Modal>
        </div>
      )
    } else {
      return (
        <section>
          Component Loading . .
        </section>
      )
    }

  }

  /*
  async bindColleagueDetails(data) {
    let client: MSGraphClient = await this.props.graphClient.getClient();
    let userList = await client.api('/users/' + data + '/people').get();
    userList.value.map(function (colleagueData) {
      console.log(colleagueData);
      // userList.value[0].personType.class
    });
    return (<div>Testing</div>);
  }
  */

  async OpenModalWithData(data) {
    let currentUserId = data.currentTarget.getAttribute("data-id");
    let currentTitle = data.currentTarget.getAttribute("data-title");
    let currentEmail = data.currentTarget.getAttribute("data-email");
    let currentDoj = data.currentTarget.getAttribute("data-doj");
    let workPhone = data.currentTarget.getAttribute("data-workPhone");
    let mobilePhone = data.currentTarget.getAttribute("data-mobilePhone");
    let officeLocation = data.currentTarget.getAttribute("data-officeLocation");
    let photoLink = data.currentTarget.getAttribute("data-photoLink");
    let department = data.currentTarget.getAttribute("data-department");
    colleagueDetails = [];
    let managerName = "";
    let managerDoj = "";
    let manaferPhotoLink = "";
    let isManagerExist = true;
    let isColleaguesExist = true;
    let client: MSGraphClient = await this.props.graphClient.getClient();
    let managerApi: any = await client.api('/users/' + currentUserId + '/manager').get()
      .catch((e) => {
        isManagerExist = false;
        managerDetails.push({
          "name": "not found",
          "doj": "not found "
        });
      });

    let colleaguesApi: any = await client.api('/users/' + currentUserId + '/directReports').get()
      .catch((e) => {
        console.log("Colleagues API not found ");
        isColleaguesExist = false;
      });

    // console.log(managerApi);
    // console.log(colleaguesApi);

    if (isManagerExist == true) {
      managerName = managerApi.displayName;
      managerDoj = managerApi.jobTitle;
      manaferPhotoLink = "/_layouts/15/userphoto.aspx?size=L&username=" +managerApi.mail+ "";
    }

    if (isColleaguesExist == true) {
      if (colleaguesApi.value.length != 0) {
        for (let i = 0; i < colleaguesApi.value.length; i++) {
          if (colleaguesApi.value[i].personType.class == "directReports")
            colleagueDetails.push({
              "displayName": colleaguesApi.value[i].displayName,
              "DOJ": colleaguesApi.value[i].jobTitle,
              "photoLink": "/_layouts/15/userphoto.aspx?size=L&username=" + colleaguesApi.value[i].userPrincipalName + ""
            });
        }
      }
    }

    this.setState({
      CurrentUserId: currentUserId,
      CurrentUserDOJ: currentDoj,
      CurrentUserEmail: currentEmail,
      CurrentUserDisplayName: currentTitle,
      CurrentUserTitle: currentTitle,
      CurrentPhotoLink: photoLink,
      CurrentMobilePhone: mobilePhone,
      CurrentOfficeLocation: officeLocation,
      CurrentWorkPhone: workPhone,
      CurrentWorkDepartment: department,
      CurrentManagerDisplayName: managerName,
      CurrentManagerDOJ: managerDoj,
      CurrentManagerPhotoLink : manaferPhotoLink,
      CurrentColleagueName: colleagueDetails,
      HideModalPop: true
    });
  }

  public closeModal() {
    let reactHandler = this;
    reactHandler.setState({ HideModalPop: false });
  }
  public bindUserLetter(userName) {
    debugger;
  }
  public bindUserData(userData) {
    return (
      <Masonry breakpointCols={11} className="my-masonry-grid" columnClassName="my-masonry-grid_column" >
        {
          userData.map(function (data) {
            var divStyle = {
              backgroundImage: 'url(/_layouts/15/userphoto.aspx?size=L&username=' + data.email + ')'
            };
            return (
              <div className="user_card tooltip"
                onClick={this.OpenModalWithData.bind(this)}
                data-id={data.id}
                data-title={data.DisplayName}
                data-email={data.email}
                data-doj={data.DOJ}
                data-photoLink={data.photoLink}
                data-workPhone={data.workNumber}
                data-mobilePhone={data.mobileNumber}
                data-officeLocation={data.OfficeLocation}
                data-department={data.department}
              >
                <span className="tooltiptext">{data.DisplayName} </span>
                <div className="user_avatar" style={divStyle} ></div>
                {/* <div className="user_detail">
                  <h3 className="user_name">{data.DisplayName}</h3>
                  <p className="user_date">{data.DOJ}</p>
                </div> */}
              </div>
            );
          }.bind(this))
        }
      </Masonry>
    );
  }
}
