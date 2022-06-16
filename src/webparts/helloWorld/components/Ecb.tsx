import * as React from 'react';
import { Layer, IconButton, IButtonProps } from 'office-ui-fabric-react';
import { ContextualMenuItemType } from 'office-ui-fabric-react/lib/ContextualMenu';
import { spfi, SPFx } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAttachmentInfo } from '@pnp/sp/attachments';
// The following are project specific components
import "@pnp/sp/lists/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItem } from "@pnp/sp/items/types";
import "@pnp/sp/webs";
import "@pnp/sp/attachments";

export interface IECBProps {
    item: any;
    context: WebPartContext;
    ListName: string;
}

export class ECB extends React.Component<IECBProps, { panelOpen: boolean, contextualItems: any[] }> {

    public constructor(props: IECBProps) {
        super(props);

        this.state = {
            panelOpen: false,
            contextualItems: []
        };
    }
    componentDidMount(): void {
        this.ShowGridData();
    }

    private async ShowGridData() {



        const sp = spfi().using(SPFx(this.props.context));
        const item: any = await sp.web.lists.getByTitle(this.props.ListName).items.getById(this.props.item.Id);
        const info: IAttachmentInfo[] = await item.attachmentFiles();
        debugger
        var contextualItems = [];
        info.forEach(element => {
            contextualItems.push({
                key: element.FileName,
                name: element.FileName,
                onClick: this.handleClick.bind(element.ServerRelativeUrl, element.ServerRelativeUrl)
            })
        });
        this.setState(
            {
                ...this.state,
                contextualItems: contextualItems
            });

    }
    public render() {
        return (
            <div >
                <IconButton id='ContextualMenuButton1'

                    text=''
                    width='30'
                    split={false}
                    iconProps={{ iconName: 'Download' }}
                    menuIconProps={{ iconName: '' }}
                    menuProps={{
                        shouldFocusOnMount: true,
                        items: this.state.contextualItems
                    }} />
            </div>
        );
    }

    private handleClick(source: string, event) {

        window.open(window.location.origin + source, '_blank', 'noopener,noreferrer');
    }
}




