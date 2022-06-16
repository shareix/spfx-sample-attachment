import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import PnPTelemetry from "@pnp/telemetry-js";
import { TextField, MaskedTextField } from '@fluentui/react/lib/TextField';
import { Stack, IStackProps, IStackStyles } from '@fluentui/react/lib/Stack';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ListItemAttachments } from '@pnp/spfx-controls-react/lib/ListItemAttachments';
import { PrimaryButton, MessageBar, MessageBarType, DetailsRow, ThemeSettingName } from 'office-ui-fabric-react';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import { IItemAddResult, DateTimeFieldFormatType, IItem } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/lists/web";
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { IAttachmentInfo } from "@pnp/sp/attachments";
import "@pnp/sp/attachments";
import { initial, values } from 'lodash';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { useState } from 'react';
import "@pnp/sp/items/get-all";
import { useEffect } from 'react';
import { ECB, IECBProps } from './Ecb';


const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 5 },
  styles: { root: { width: 300 } },
};

const columnProps2: Partial<IStackProps> = {
  tokens: { childrenGap: 5 },
  styles: {
    root: { width: 300 }
  }
}
//interface||state
interface IHelloWorldState {
  ItemId: number,
  showattachment: boolean,
  CustomerName: string,
  orderid: number,
  User: [],
  AllItems: any[],
  IViewField: any[],
  GroupField: any[],
  listItemAttachmentsComponentReference: any
}


export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {

  constructor(props) {
    super(props);
  }
  state = {
    showattachment: true,
    ItemId: 0,
    CustomerName: "",
    orderid: 0,
    User: undefined,
    AllItems: undefined,
    IViewField: [{
      name: "Title",
      displayName: "MyTitle",
      linkPropertyName: "c",
      isResizable: true,
      sorting: true,
      minWidth: 0,
      maxWidth: 150,
      render: (item: any) => {
        return <a href={`${this.props.context.pageContext.site.absoluteUrl}/Lists/${this.props.ListName}/DispForm.aspx?ID=${item['ID']}`} target='_blank'>{item['Title']}</a>;
      }
    }, {
      name: "OrderNumber",
      displayName: "OrderNumber",
      linkPropertyName: "c",
      isResizable: true,
      sorting: true,
      minWidth: 0,
      maxWidth: 100

    },
    {
      name: "Files",
      sorting: false,
      maxWidth: 40,
      render: (rowitem: any) => {
        if (rowitem["Attachments"]) {
          const element: React.ReactElement<IECBProps> = React.createElement(
            ECB,
            {
              item: rowitem,
              context: this.props.context,
              ListName: this.props.ListName
            }
          );
          return element;
        }
        return <></>
      }
    }, {
      name: "StatusWorkflow",
      displayName: "StatusWorkflow",
      linkPropertyName: "c",
      isResizable: true,
      sorting: true,
      minWidth: 0,
      maxWidth: 100,

      render: (item: any) => {
        if (item['StatusWorkflow'] == "Approve") {
          return <div className={'Approve-' + item['Id']} style={{ color: 'black', width: '100%', height: '100%', position: 'absolute', left: '0', top: '0', backgroundColor: 'rgba(186, 216, 10, 0.2)', paddingLeft: '10px', paddingTop: '10px' }}>{item['StatusWorkflow']}
            <style>
              {/* {'.Approve-' + item['Id'] + '> .ms-DetailsRow-cell { background-color: red;}'} */}
            </style>
          </div>;
        }
        else {
          return <h4 style={{ color: "black", backgroundColor: 'rgba(232, 17, 35, 0.1)', width: '100%', height: '100%', position: 'absolute', left: '0', top: '-15px', fontWeight: '400', paddingLeft: '10px', paddingTop: '10px' }}>{item['StatusWorkflow']}
            <style>
              {'.ms-FocusZone div:first-child div.ms-DetailsRow-cell:last-child {background-color: rgba(232, 17, 35, 0.1);}'}
            </style>
          </h4>;
        }
      }

    }],
    GroupField: undefined,
    listItemAttachmentsComponentReference: React.createRef<ListItemAttachments>()

  };

  componentDidMount(): void {
    this.ShowGridData();
  }

  private async ShowGridData() {

    debugger
    const sp = spfi().using(SPFx(this.props.context));
    const allItems: any[] = await sp.web.lists.getByTitle(this.props.ListName).items.getAll();
    console.log("allItems:", allItems);
    console.log("allItems.length:", allItems.length);
    this.setState({ AllItems: allItems });
  }

  private async clear() {
    debugger
    this.setState({

      ...this.state,
      showattachment: (!this.state.showattachment),
      ItemId: 0,
      listItemAttachmentsComponentReference: React.createRef<ListItemAttachments>(),
      CustomerName: "",
      orderid: 0,
      User: undefined,
    });

  }
  private async createNewItem(refattachment) {
    debugger
    var cc = this.props.context;
    const groupByFields: IGrouping[] = [
      {
        name: "Status",
        order: GroupOrder.ascending
      },];
    this.setState({ GroupField: groupByFields })
    console.log("groupByFields:", groupByFields)
    console.log("GroupField:", this.state.GroupField)

    const sp = spfi().using(SPFx(this.props.context));

    const iar: IItemAddResult = await sp.web.lists.getByTitle(this.props.ListName).items.add({
      Title: this.state.CustomerName,
      OrderNumber: this.state.orderid,
      StatusWorkflow: 'Pending',
    });
    const items2: any[] = await sp.web.lists.getByTitle(this.props.ListName).items.select("Title")();
    console.log("items2", items2);

    await refattachment.current.uploadAttachments(iar.data.ID);

    const result = await sp.web.lists.getByTitle(this.props.ListName).items.getById(iar.data.ID).validateUpdateListItem([{
      FieldName: "User",
      FieldValue: JSON.stringify([{ "Key": `${this.state.User}` }]),
    }]);
    debugger
    this.ShowGridData();
    this.setState({


      showattachment: false,
    });

    this.setState({
      ...this.state,
      showattachment: true,
      ItemId: 0,
      listItemAttachmentsComponentReference: React.createRef<ListItemAttachments>(),
      CustomerName: "",
      orderid: 0,
      User: undefined,
    });
  }



  private onchange(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string, IHelloWorldState): void {
  }



  private handleChangeCustomerName = (event) => {

    const CustomerName = event.target.value
    this.setState({ CustomerName: CustomerName })
    return CustomerName
  }

  private handleChangeorderid = (event) => {

    const orderid = event.target.value
    this.setState({ orderid: orderid })
    return orderid
  }
  private handleChangeUser = (event) => {

    const User = event[0].loginName
    this.setState({ User: User })
    return User
  }


  private _getSelection(items: any[]) {
    console.log('Selected items:', items);
  }




  public render(): React.ReactElement<IHelloWorldProps> {
    const telemetry = PnPTelemetry.getInstance();
    telemetry.optOut();

    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;


    const context = this.props.context;



    return (
      <>
        <section className={`${styles.helloWorld} ${hasTeamsContext ? styles.teams : ''}`}>
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              <TextField onChange={this.handleChangeorderid} value={this.state.orderid.toString()} name='orderid' type='number' label="Order Number" borderless placeholder="No Order Number." />
              <TextField onChange={this.handleChangeCustomerName} value={this.state.CustomerName.toString()} name='CustomerName' label="Customer Name" borderless placeholder="No Customer Name." />
              <TextField label="Destination" borderless placeholder="No Destination." />
              <PeoplePicker onChange={this.handleChangeUser} defaultSelectedUsers={this.state.User} titleText='Owner'
                context={this.props.context}
                showtooltip={true}
                disabled={false}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={500} />

            </Stack>
            <Stack {...columnProps}>
              {this.state.showattachment == true &&
                <ListItemAttachments listId='d09b1003-8a2b-420d-9d07-f54323841909'
                  ref={this.state.listItemAttachmentsComponentReference}
                  itemId={this.state.ItemId}
                  context={this.props.context}
                  label='Attachment'
                  disabled={false}

                />
              }
            </Stack>
          </Stack>
          <PrimaryButton text="Create New Item" onClick={() => this.createNewItem(this.state.listItemAttachmentsComponentReference)} />
          <br />
          <br />
          <ListView
            items={this.state.AllItems}
            viewFields={this.state.IViewField}
            iconFieldName="MoreVertical"
            showFilter={true}
            groupByFields={this.state.GroupField}
          />
        </section>
      </>
    );
  }

}
