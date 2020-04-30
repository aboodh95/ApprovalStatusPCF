import * as React from "react";
import * as moment from "moment";
import { initializeIcons } from "@uifabric/icons";
import {
  Text,
  Stack,
  IStackTokens,
  FontWeights,
  ITextStyles,
  IPersonaSharedProps,
  Persona,
  PersonaSize,
  IconButton,
  IIconProps,
  Spinner,
  SpinnerSize,
} from "office-ui-fabric-react";

import { IInputs } from "../generated/ManifestTypes";
import { FlowLoader } from "../api/FlowLoader";
import { SendReminder } from "../api/SendReminder";

export interface IFlowDetailStatus {
  isLoading: boolean;
  flow?: IFlow;
  message: string;
}

export interface IFlowDetailProps {
  Context: ComponentFramework.Context<IInputs>;
}

export interface IFlow {
  name: string;
  startedOn: Date;
  users: IPersonaSharedProps[];
  entityReference: { id: string; entityName: string };
}

export class FlowDetail extends React.Component<
  IFlowDetailProps,
  IFlowDetailStatus
> {
  constructor(prop: IFlowDetailProps) {
    super(prop);
    initializeIcons();
    this.state = {
      isLoading: true,
      flow: undefined,
      message: "There is no running flow associated with this record",
    };
    let flowLoader = new FlowLoader(this.props.Context.webAPI);
    flowLoader.loadFlowStatus((state) => {
      this.setState(state);
    });
  }

  render() {
    let width =
      document.getElementsByClassName(
        "customControl hamwi ApprovalStatus hamwi.ApprovalStatus"
      )[0].clientWidth * 0.95;
    const stackToken: IStackTokens = {
      childrenGap: 10,
      padding: 1,
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
    const sendEmailIcon: IIconProps = { iconName: "Mail" };

    return this.state.isLoading ? (
      <Spinner size={SpinnerSize.large} />
    ) : this.state.flow ? (
      <Stack tokens={stackToken}>
        <Stack.Item>
          <Text styles={flowNameStyle}>Flow: {this.state.flow?.name}</Text>
        </Stack.Item>
        <Stack.Item>
          <Text variant="small" styles={flowDateStyle}>
            Started On:{" "}
            {moment(this.state.flow?.startedOn).format("YYYY-MM-DD HH:mm:ss")}
          </Text>
        </Stack.Item>
        {this.state.flow
          ? this.state.flow.users.map((user) => {
              return (
                <Stack.Item key={user.id}>
                  <Stack horizontal>
                    <Stack.Item grow>
                      <Persona
                        key={user.id}
                        {...user}
                        size={PersonaSize.size40}
                      />
                    </Stack.Item>
                    <Stack.Item>
                      <IconButton
                        iconProps={sendEmailIcon}
                        title="Send Reminder Email"
                        ariaLabel="Send Reminder Email"
                        checked={true}
                        onClick={() => {
                          if (this.state.flow && user.id) {
                            SendReminder(
                              this.state.flow,
                              user.id,
                              this.props.Context
                            );
                          }
                        }}
                      />
                    </Stack.Item>
                  </Stack>
                </Stack.Item>
              );
            })
          : ""}
      </Stack>
    ) : (
      <Text styles={flowNameStyle}>{this.state.message}</Text>
    );
  }
}
