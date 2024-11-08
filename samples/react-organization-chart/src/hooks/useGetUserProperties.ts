
import { sp, SPBatch, Web} from "@pnp/sp/";
import { IUserInfo } from "../models/IUserInfo";
import * as React from "react";
import { get, set } from "idb-keyval";
import { sortBy, filter } from "lodash";
import { IPersonProperties } from "../models/IPersonProperties";
import { getGUID } from "@pnp/common";

/*************************************************************************************/
// Hook to get user profile information
// *************************************************************************************/

type getUserProfileFunc = ( currentUser: string,
  startUser?: string,
  showAllManagers?: boolean) => Promise<returnProfileData>;

type returnProfileData =  { managersList:IUserInfo[], reportsLists:IUserInfo[], currentUserProfile :IPersonProperties} ;

export const useGetUserProperties  =  ():  { getUserProfile:getUserProfileFunc }  => {

  const getUserProfile = React.useCallback(
    async (
      currentUser: string,
      startUser?: string,
      showAllManagers?: boolean
    ): Promise<returnProfileData> => {
      if (!currentUser) return;
      const loginName = currentUser;
      const loginNameStartUser: string = startUser && startUser;
      const cacheCurrentUser:IPersonProperties = await get(`${loginName}__orgchart__`);
      let currentUserProfile:IPersonProperties   = undefined;
      if (!cacheCurrentUser) {
        currentUserProfile = await sp.profiles.getPropertiesFor(loginName);
        // TODO: Get Extended properties from list
        currentUserProfile = await getExtendedProfileData(currentUserProfile);
        await set(`${loginName}__orgchart__`, currentUserProfile);
      } else {
        currentUserProfile = cacheCurrentUser;
      }
      // get Managers and Direct Reports
      let reportsLists: IUserInfo[] = [];
      let openPositionsLists: IUserInfo[] = [];
      let managersList: IUserInfo[] = [];

      const wDirectReports: string[] =
        currentUserProfile && currentUserProfile.DirectReports;
      const wExtendedManagers: string[] =
        currentUserProfile && currentUserProfile.ExtendedManagers;

      // Get Direct Reports if exists
      if (wDirectReports && wDirectReports.length > 0) {
        reportsLists = await getDirectReports(wDirectReports);
      }

      openPositionsLists = await getOpenPositions(currentUserProfile)
      if(openPositionsLists){
        reportsLists = reportsLists.concat(openPositionsLists);
      }

      // Get Managers if exists
      if (startUser && wExtendedManagers && wExtendedManagers.length > 0) {
        managersList = await getExtendedManagers(
          wExtendedManagers,
          loginNameStartUser,
          showAllManagers
        );
      }

      return   { managersList, reportsLists, currentUserProfile } ;
    },
    []
  );

  return   { getUserProfile }  ;
};

const getDirectReports = async (
  directReports: string[]
): Promise<IUserInfo[]> => {
  const _reportsList: IUserInfo[] = [];
  const batch: SPBatch = sp.createBatch();
  for (const userReport of directReports) {
    const cacheDirectReport: IPersonProperties = await get(`${userReport}__orgchart__`);
    if (!cacheDirectReport) {
      sp.profiles
        .inBatch(batch)
        .getPropertiesFor(userReport)
        .then(async (directReport: IPersonProperties) => {
          await getExtendedProfileData(directReport).then(async (directReport) => {
            _reportsList.push(await manpingUserProperties(directReport));
            await set(`${userReport}__orgchart__`, directReport);

          });

        });
        // TODO: Get Extended properties from list

    } else {
      _reportsList.push(await manpingUserProperties(cacheDirectReport));
    }
  }
  await batch.execute();
  return sortBy(_reportsList, ["displayName"]);
};

const getOpenPositions = async (
  user: IPersonProperties,
): Promise<IUserInfo[]> => {
  const _openPositionsList: IUserInfo[] = [];
  const openPositionsResponse = await sp.web.lists.getByTitle("OpenPositions").items
  .select("Title", "Description","Owner/UserName","Owner/EMail")
  .expand("Owner")
  .filter("Owner/EMail eq '" + user.Email + "'")
  .orderBy("Modified", true).get();
  if(openPositionsResponse.length > 0){
    for (const openPosition of openPositionsResponse) {
      const openPositionObj: IPersonProperties = {
        DirectReports: [],
        AccountName : openPosition.Title,
        DisplayName : openPosition.Title,
        Title : openPosition.Description,
        UserProfileProperties : user.UserProfileProperties,
        Email: user.Email,
        ExtendedManagers: null,
        ExtendedReports:null,
        IsFollowed: false,
        LatestPost: null,
        Peers: [],
        PersonalSiteHostUrl: null,
        PersonalUrl: null,
        PictureUrl: null,
        UserUrl: null,
        loginName: null,
        Skills: null,
        Certifications: null,
        SocialNetwork: null
      };
      const openPositionObjMp = await manpingUserProperties(openPositionObj);
      openPositionObjMp.manager = user.AccountName;
      _openPositionsList.push(openPositionObjMp);
    }
    return sortBy(_openPositionsList, ["displayName"]);
  }
  return null;
};


const getHasOpenPositions = async (
  user: IPersonProperties,
): Promise<boolean> => {
  const openPositionsResponse = await sp.web.lists.getByTitle("OpenPositions").items
  .select("Title", "Description","Owner/UserName","Owner/EMail")
  .expand("Owner")
  .filter("Owner/EMail eq '" + user.Email + "'")
  .orderBy("Modified", true).get();
  if(openPositionsResponse.length > 0){
    
    return true;
  }
  return false;
};

const getExtendedManagers = async (
  extendedManagers: string[],
  startUser: string,
  showAllManagers: boolean
): Promise<IUserInfo[]> => {
  const wManagers: IUserInfo[] = [];
  const batch: SPBatch = sp.createBatch();

  for (const manager of extendedManagers) {
    if (!showAllManagers && manager !== startUser) {
      continue;
    }
    const cacheManager: IPersonProperties = await get(`${manager}__orgchart__`);
    if (!cacheManager) {
      sp.profiles
        .inBatch(batch)
        .getPropertiesFor(manager)        
        .then(async (_profile: IPersonProperties) => {
          await getExtendedProfileData(_profile).then(async (_profile) => {
            wManagers.push(await manpingUserProperties(_profile));
            await set(`${manager}__orgchart__`, _profile);

          });
        });
        // TODO: Get Extended properties from list

    } else {
      wManagers.push(await manpingUserProperties(cacheManager));
    }
  }
  await batch.execute();
  return wManagers;
};

const getExtendedProfileData = async(user: IPersonProperties) : Promise<IPersonProperties> => {
  const additionalInformation = await sp.web.lists.getByTitle("AdditionalProfileInformation").items
  .select("Title", "Skills", "LinkedIn","User/UserName","User/EMail", "Certifications/Issuer", "Certifications/Title")
  .expand("User,Certifications")
  .filter("User/EMail eq '" + user.Email + "'")
  .top(1).orderBy("Modified", true).get();
  if(additionalInformation.length > 0){
    user.Skills = additionalInformation[0].Skills;
    user.Certifications = additionalInformation[0].Certifications;
    if(additionalInformation[0].LinkedIn){
      user.SocialNetwork = [];
      user.SocialNetwork.push("<a href='" + additionalInformation[0].LinkedIn.Url + "'>" + additionalInformation[0].LinkedIn.Description + "</a>");
    }

  }
  return user;
}

export const manpingUserProperties = async (
  userProperties: IPersonProperties
): Promise<IUserInfo> => {
  const hasDirectReportsVar = await getHasOpenPositions(userProperties);
  return {
    displayName: userProperties.DisplayName as string,
    email: userProperties.Email as string,
    title: userProperties.Title as string,
    pictureUrl: userProperties.PictureUrl,
    id: userProperties.AccountName,
    userUrl: userProperties.UserUrl,
    numberDirectReports: userProperties.DirectReports.length,
    hasDirectReports: userProperties.DirectReports.length > 0 ? true : hasDirectReportsVar,
    hasPeers: userProperties.Peers.length > 0 ? true : false,
    numberPeers: userProperties.Peers.length,
    department: filter(userProperties?.UserProfileProperties,{"Key": "Department"})[0].Value ?? '',
    workPhone: filter(userProperties?.UserProfileProperties,{"Key": "WorkPhone"})[0].Value ?? '',
    cellPhone: filter(userProperties?.UserProfileProperties,{"Key": "CellPhone"})[0].Value ?? '',
    location: filter(userProperties?.UserProfileProperties,{"Key": "SPS-Location"})[0].Value ?? '',
    office: filter(userProperties?.UserProfileProperties,{"Key": "Office"})[0].Value ?? '',
    manager: filter(userProperties?.UserProfileProperties,{"Key": "Manager"})[0].Value ?? '',
    loginName: userProperties.loginName,
    skills: userProperties.Skills?.join(", ") ?? null,
    certifications: userProperties.Certifications?.map((x) => {return x.Title + " from " + x.Issuer}).join(", ") ?? null,
    socialnetwork: userProperties.SocialNetwork?.join() ?? null
  };
};
