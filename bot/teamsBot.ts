import {
  TeamsActivityHandler,
  TurnContext,
  TaskModuleContinueResponse,
  MessagingExtensionAction,
  MessagingExtensionActionResponse,
} from "botbuilder";

export interface DataInterface {
  likeCount: number;
}

export class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
  }

  public async handleTeamsMessagingExtensionSubmitAction(_context: TurnContext, 
    _action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
      
      if (_action.data && _action.data.status === "completed") {
        return {
          task: {
              type: "message",
              value: "Thanks!"
          }
        };
      }
      else {
        return {
          task: {
            type: "continue",
            value: {
              url: "https://localhost:53000/#/tab",
              title: "Scan report",
              height: "large",
              width: "large" 
            }
          } as TaskModuleContinueResponse
        }
      }
  }
}
