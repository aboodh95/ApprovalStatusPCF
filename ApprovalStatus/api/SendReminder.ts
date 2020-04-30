import { IFlow } from "../components/FlowDetail";
import { IInputs } from "../generated/ManifestTypes";

export const SendReminder = async (
  flow: IFlow,
  sendToUserId: string,
  context: ComponentFramework.Context<IInputs>
) => {
  try {
    var email: ComponentFramework.WebApi.Entity = {};
    var entityMetadata = await context.utils.getEntityMetadata(
      flow.entityReference.entityName
    );
    var primaryNameAttribute = entityMetadata["_primaryNameAttribute"];
    var entity = await context.webAPI.retrieveRecord(
      flow.entityReference.entityName,
      flow.entityReference.id,
      `?$select=${primaryNameAttribute}`
    );
    var primaryNameValue = entity[primaryNameAttribute]
      ? entity[primaryNameAttribute]
      : "(No Name)";
    var regarding = `regardingobjectid_${flow.entityReference.entityName}@odata.bind`;
    email[
      regarding
    ] = `/${entityMetadata["_entitySetName"]}(${flow.entityReference.id})`;
    email[`description`] = `Hi,
  <br />
  <br />
  A kind reminder to finalize the ${flow.name} that's related to: ${primaryNameValue}.
  <br />
  <br />
  Thanks `;
    email[`subject`] = `Reminder for ${flow.name}`;
    email[`email_activity_parties`] = [
      {
        "partyid_systemuser@odata.bind": `/systemusers(${context.userSettings.userId
          .replace("{", "")
          .replace("}", "")})`,
        participationtypemask: 1,
      },
      {
        "partyid_systemuser@odata.bind": `/systemusers(${sendToUserId})`,
        participationtypemask: 2,
      },
    ];

    var emailId = await context.webAPI.createRecord("email", email);
    context.navigation.openForm({
      entityName: "email",
      entityId: emailId.id,
      openInNewWindow: true,
    });
  } catch (error) {
    console.error(error);
    alert(error.message);
  }
};
