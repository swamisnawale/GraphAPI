import * as React from "react";
// import styles from "./GraphApi.module.scss";
import { IGraphApiProps } from "./IGraphApiProps";
import { MSGraphClientV3 } from "@microsoft/sp-http";

// Email Interface
interface IEmails {
  subject: string;
  webLink: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  receivedDateTime: any;
  bodyPreview: string;
  isRead: any;
}

// All Items Interface
interface IAllItems {
  AllEmails: IEmails[];
}
export default class GraphApi extends React.Component<
  IGraphApiProps,
  IAllItems
> {
  constructor(props: IGraphApiProps, state: IAllItems) {
    super(props);
    this.state = {
      AllEmails: [],
    };
  }

  componentDidMount(): void {
    this.getMyEmails();
    this.getMyProfile();
  }

  public getMyProfile() {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("/me")
          .version("v1.0")
          .get((err: any, res: any) => {
            console.log(res);
            // console.log(err);
          });
      });
  }
  public getMyEmails() {
    this.props.context.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3): void => {
        client
          .api("/me/messages")
          .version("v1.0")
          .select("subject,webLink,from,receivedDateTime,isRead,bodyPreview")
          .get((err: any, res: any) => {
            this.setState({
              AllEmails: res.value,
            });
            // console.log(this.state.AllEmails);
            // console.log(res);
            // console.log(err);
          });
      });
  }

  public render(): React.ReactElement<IGraphApiProps> {
    return (
      <div>
        {this.state.AllEmails.map((email) => {
          return (
            <div>
              <p>{email.from.emailAddress.name}</p>
              <p>{email.subject}</p>
              <p>{email.receivedDateTime}</p>
              <p>{email.bodyPreview}</p>
              <button
                onClick={() => {
                  window.open(email.webLink, "_blank");
                }}
              >
                {" "}
                open email in new tab
              </button>
              <hr />
            </div>
          );
        })}
      </div>
    );
  }
}
