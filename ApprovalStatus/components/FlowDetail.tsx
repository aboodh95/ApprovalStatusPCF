import * as React from "react";
import * as moment from "moment";

import { Announced } from "office-ui-fabric-react/lib/Announced";
import {
  Text,
  Stack,
  IStackTokens,
  Button,
  IStackItemTokens,
  FontWeights,
  ITextStyles,
} from "office-ui-fabric-react";
import { Card, ICardTokens } from "@uifabric/react-cards";
import {
  IPersonaSharedProps,
  Persona,
  PersonaSize,
} from "office-ui-fabric-react/lib/Persona";
import { IInputs } from "../generated/ManifestTypes";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";

interface IFlowDetailStatus {
  isLoading: boolean;
  flow?: IFlow;
}

export interface IFlowDetailProps {
  Context: ComponentFramework.Context<IInputs>;
}

export interface IFlow {
  name: string;
  startedOn: Date;
  users: IPersonaSharedProps[];
}

export class FlowDetail extends React.Component<
  IFlowDetailProps,
  IFlowDetailStatus
> {
  constructor(prop: IFlowDetailProps) {
    super(prop);
    this.state = {
      isLoading: true,
      flow: undefined,
    };
    this.loadFlowStatus();
  }

  render() {
    let width =
      document.getElementsByClassName(
        "customControl hamwi ApprovalStatus hamwi.ApprovalStatus"
      )[0].clientWidth * 0.8;
    const cardTokens: ICardTokens = {
      childrenMargin: 12,
      width: width,
      minWidth: width,
    };
    const flowNameStyle: ITextStyles = {
      root: {
        color: "#333333",
        fontWeight: FontWeights.semibold,
      },
    };
    const flowDateStyle: ITextStyles = {
      root: {
        color: "#333333",
        fontWeight: FontWeights.regular,
      },
    };
    return this.state.isLoading ? (
      <Spinner size={SpinnerSize.large} />
    ) : this.state.flow ? (
      <Card tokens={cardTokens}>
        <Card.Section>
          <Text styles={flowNameStyle}>Flow Name: {this.state.flow?.name}</Text>
          <Text variant="small" styles={flowDateStyle}>
            Started On:{" "}
            {moment(this.state.flow?.startedOn).format("YYYY-MM-DD HH:mm:ss")}
          </Text>
        </Card.Section>
        <Card.Section>
          <Stack tokens={{ childrenGap: 10 }}>
            {this.state.flow
              ? this.state.flow.users.map((user) => {
                  return (
                    <Stack.Item key={user.imageInitials}>
                      <Persona
                        key={user.id}
                        {...user}
                        hidePersonaDetails={false}
                        size={PersonaSize.size32}
                      />
                    </Stack.Item>
                  );
                })
              : ""}
          </Stack>
        </Card.Section>
      </Card>
    ) : (
      <Card tokens={cardTokens}>
        <Card.Item>
          <Text styles={flowNameStyle}>
            There is no running flow associated with this record
          </Text>
        </Card.Item>
      </Card>
    );
  }

  async loadFlowStatus() {
    var params = this.getPageParameters();
    var approvalFetchData = {
      msdyn_flow_approval_itemlink: `%${params.id}%`,
      statecode: "0",
    };
    //Load flow Approval records
    try {
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
      var webApi = this.props.Context.webAPI;
      approvalFetchXml = "?fetchXml=" + encodeURIComponent(approvalFetchXml);
      var result = await webApi.retrieveMultipleRecords(
        "msdyn_flow_approval",
        approvalFetchXml
      );
      if (result.entities.length == 0) {
        this.setState({
          isLoading: false,
        });
        return;
      }
      var flowEntity = result.entities[0];
      let flow: IFlow = {
        name: flowEntity["msdyn_flow_approval_title"],
        startedOn: new Date(flowEntity["createdon"]),
        users: [],
      };

      //Load flow Approval response records
      var requestFetchData = {
        statecode: "0",
        msdyn_flow_approvalrequest_approval:
          flowEntity["msdyn_flow_approvalid"],
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
      result = await webApi.retrieveMultipleRecords(
        "msdyn_flow_approvalrequest",
        requestFetchXml
      );

      if (result.entities.length != 0) {
        var userFetchXml = [
          "<fetch>",
          "  <entity name='systemuser'>",
          "    <attribute name='entityimage_url' />",
          "    <attribute name='systemuserid' />",
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
        result = await webApi.retrieveMultipleRecords(
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
          });
        }
      }
      console.log(flow);
      this.setState({
        isLoading: false,
        flow: flow,
      });
    } catch (error) {
      console.error(error);
    }
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
