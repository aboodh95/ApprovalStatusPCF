import { IFlowDetailStatus } from "../components/FlowDetail";
import { IFlow } from "../components/FlowDetail";

export class FlowLoader {
  webApi: ComponentFramework.WebApi;
  recordId: string;
  entityName: string;
  appId: string;
  constructor(webApi: ComponentFramework.WebApi) {
    this.webApi = webApi;
    var params = this.getPageParameters();
    this.recordId = params.id;
    this.entityName = params.etn;
    this.appId = params.appid;
  }
  /**
   *
   */

  async loadFlowStatus(callBack: (state: IFlowDetailStatus) => void) {
    var approvalFetchData = {
      msdyn_flow_approval_itemlink: `%${this.recordId}%`,
      statecode: "0",
    };
    //Load flow Approval records
    try {
      var result = await this.getRunningFlows(approvalFetchData);
      if (result.entities.length == 0) {
        callBack({
          isLoading: false,
          message: "No running flows related to this record",
        });
        return;
      }
      let flow: IFlow | undefined;
      ({ flow, result } = await this.getFlowResponse(result));

      result = await this.getUsers(result, flow);
      callBack({
        isLoading: false,
        flow: flow,
        message: "",
      });
    } catch (error) {
      let errorMessage = `an error just occurred: ${error.message}`;
      if (
        error.message
          .toString()
          .indexOf(
            "You do not have {0} permission to access {1} records. Contact your Microsoft Dynamics 365 administrator for help."
          ) > -1
      ) {
        errorMessage =
          "You don't have permission for the Approval Entity, Please check with IT Administrator";
      }
      callBack({
        message: errorMessage,
        isLoading: false,
      });
      console.error(error);
    }
  }

  private async getUsers(
    result: ComponentFramework.WebApi.RetrieveMultipleResponse,
    flow: IFlow
  ) {
    if (result.entities.length != 0) {
      var userFetchXml = [
        "<fetch>",
        "  <entity name='systemuser'>",
        "    <attribute name='entityimage_url' />",
        "    <attribute name='systemuserid' />",
        "    <attribute name='domainname' />",
        "    <attribute name='fullname' />",
        "    <filter>",
        "      <condition attribute='systemuserid' operator='in'>",
      ];
      for (let index = 0; index < result.entities.length; index++) {
        const element = result.entities[index];
        userFetchXml.push(
          `<value>${element["msdyn_flow_approvalrequestidx_owninguserid"]}</value>`
        );
      }
      userFetchXml.push(
        "      </condition>",
        "    </filter>",
        "  </entity>",
        "</fetch>"
      );
      result = await this.webApi.retrieveMultipleRecords(
        "systemuser",
        "?fetchXml=" + encodeURIComponent(userFetchXml.join(""))
      );
      for (let index = 0; index < result.entities.length; index++) {
        const element = result.entities[index];
        flow.users.push({
          id: element["systemuserid"],
          imageUrl: element["entityimage_url"],
          presenceTitle: element["fullname"],
          text: element["fullname"],
          secondaryText: element["domainname"],
        });
      }
    }
    return result;
  }

  private async getFlowResponse(
    result: ComponentFramework.WebApi.RetrieveMultipleResponse
  ) {
    var flowEntity = result.entities[0];
    let flow: IFlow = {
      name: flowEntity["msdyn_flow_approval_title"],
      startedOn: new Date(flowEntity["createdon"]),
      users: [],
      entityReference: {
        id: this.recordId,
        entityName: this.entityName,
      },
    };
    //Load flow Approval response records
    var requestFetchData = {
      statecode: "0",
      msdyn_flow_approvalrequest_approval: flowEntity["msdyn_flow_approvalid"],
    };
    var requestFetchXml = [
      "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>",
      "  <entity name='msdyn_flow_approvalrequest'>",
      "    <attribute name='msdyn_flow_approvalrequestid' />",
      "    <attribute name='msdyn_flow_approvalrequestidx_owninguserid' />",
      "    <order attribute='msdyn_flow_approvalrequest_name' descending='false' />",
      "    <filter type='and'>",
      "      <condition attribute='statecode' operator='eq' value='",
      requestFetchData.statecode /*0*/,
      "'/>",
      "      <condition attribute='msdyn_flow_approvalrequest_approval' operator='eq' value='",
      requestFetchData.msdyn_flow_approvalrequest_approval,
      "'/>",
      "    </filter>",
      "  </entity>",
      "</fetch>",
    ].join("");
    requestFetchXml = "?fetchXml=" + encodeURIComponent(requestFetchXml);
    result = await this.webApi.retrieveMultipleRecords(
      "msdyn_flow_approvalrequest",
      requestFetchXml
    );
    return { flow, result };
  }

  private async getRunningFlows(approvalFetchData: {
    msdyn_flow_approval_itemlink: string;
    statecode: string;
  }) {
    var approvalFetchXml = [
      "<fetch top='1'>",
      "  <entity name='msdyn_flow_approval'>",
      "    <attribute name='msdyn_flow_approvalid' />",
      "    <attribute name='msdyn_flow_approval_title' />",
      "    <attribute name='createdon' />",
      "    <filter>",
      "      <condition attribute='statecode' operator='eq' value='",
      approvalFetchData.statecode,
      "'/>",
      "      <condition attribute='msdyn_flow_approval_itemlink' operator='like' value='",
      approvalFetchData.msdyn_flow_approval_itemlink,
      "'/>",
      "    </filter>",
      "    <order attribute='createdon' descending='true' />",
      "  </entity>",
      "</fetch>",
    ].join("");
    approvalFetchXml = "?fetchXml=" + encodeURIComponent(approvalFetchXml);
    var result = await this.webApi.retrieveMultipleRecords(
      "msdyn_flow_approval",
      approvalFetchXml
    );
    return result;
  }

  getPageParameters(): {
    appid: string;
    pagetype: string;
    etn: string;
    id: string;
  } {
    const url = window.location.href;
    const parametersString = url.split("?")[1];
    let parametersObj: any = {};
    if (parametersString) {
      for (let paramPairStr of parametersString.split("&")) {
        let paramPair = paramPairStr.split("=");
        parametersObj[paramPair[0]] = paramPair[1];
      }
    }
    return parametersObj;
  }
}
