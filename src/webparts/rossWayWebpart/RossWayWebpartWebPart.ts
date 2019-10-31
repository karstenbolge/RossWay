import { Version } from '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    ISerializedServerProcessedData,
    PropertyPaneDropdownOptionType,
    PropertyPaneDropdown,
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './RossWayWebpartWebPart.module.scss';
import * as strings from 'RossWayWebpartWebPartStrings';
import MockHttpClient from './MockHttpClient';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';

export interface IRossWayWebpartWebPartProps {
    description: string;
    project: string;
}

export interface ISPLists {
    value: ISPList[];
}

export interface ISPList {
    Title: string;
    Id: string;
}

export interface ISPListDocuments {
    value: ISPListDocument[];
}

export interface ISPListDocument {
    Title: string;
    Id: string;
    Deliverables: string;
    Phase: string;
    Status1: string;
    ServerRedirectedEmbedUrl: string;
}

export interface ISPListDatas {
    value: ISPListData[];
}

export interface ISPListData {
    Title: string;
    Status: string;
    Id: string;
}

interface ISPProject {
    FileRef: string;
    Id: string;
}

export default class RossWayWebpartWebPart extends BaseClientSideWebPart<IRossWayWebpartWebPartProps> {
    private projectsFetched: boolean;
    private projectsOptions: IPropertyPaneDropdownOption[];

    private documentsGuidFetched: boolean;
    private documentsGuid: string;

    private fetchLists(url: string): Promise<any> {
        return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            });
    }

    private async fetchProjects(): Promise<IPropertyPaneDropdownOption[]> {
        const response = await this.fetchLists(this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists(guid'b28e0d0a-548b-4fbc-95f8-fac3b3b44029')/Items?$select=*,EncodedAbsUrl,FileRef,Id,Deliverables,Phase,Status1&$filter=FSObjType eq 1");
        var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        response.value.map((list: ISPProject) => {
            options.push({ key: list.FileRef.substr(list.FileRef.lastIndexOf('/') + 1), text: list.FileRef.substr(list.FileRef.lastIndexOf('/') + 1) });
        });
        return options;
    }

    private fetchDocumentsGuidAsync(): void {
        if (Environment.type === EnvironmentType.Local) {
            this.documentsGuid = " GUID";
            return;
        }

        this.fetchLists(this.context.pageContext.web.absoluteUrl + "/_api/web/lists").
            then((response) => {
                this.documentsGuidFetched = true;
                this.documentsGuid = response.value.filter((item) => {
                    if (item.Title === "Documents") return true;
                })[0].Id;

                this.fetchLists(this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists(guid'b28e0d0a-548b-4fbc-95f8-fac3b3b44029')/Items?$select=FSObjType,EncodedAbsUrl,FileRef,Id,Deliverables,Phase,Status1&$filter=startswith(FileRef, '/sites/RossManagement/Delte dokumenter/" + this.properties.project + "/')")// /Projekt 11 Demo')")
                    .then((response2) => {
                        console.log("Svaret er " + JSON.stringify(response2));
                        //return response.json();

                        let html: string = "Docuements <TABLE><TR><TD>Id</TD><TD>Title</TD><TD>Deliverables</TD><TD>Phase</TD><TD>Status1</TD><TD>Link</TD></TR>";
                        response2.value.map((list) => {
                            if (list.FSObjType === 0) {
                                html += `<TR><TD>${list.Id}</TD><TD>${list.Title}</TD><TD>${list.Deliverables}</TD><TD>${list.Phase}</TD><TD>${list.Status1}</TD><TD><a href="${list.EncodedAbsUrl}">${list.FileRef.substr(list.FileRef.lastIndexOf('/') + 1)}</a></TD></TR>`;
                            }
                        });
                        html += "</TABLE>";



                        const listContainer: Element = this.domElement.querySelector("#spListItemContainer");
                        listContainer.innerHTML = html;
                    });
            });
    }

    private _renderListAsync(): void {
        if (Environment.type === EnvironmentType.Local) {
            this._getMockListData().then((response) => {
                this._renderList(response.value);
            });
            return;
        }

        this._getListData().
            then((response) => {
                console.log("Karsten json ", response);
                this._renderList(response.value);
            });
    }

    private _renderListDocumentAsync(): void {
        if (Environment.type === EnvironmentType.Local) {
            this._getMockListDocument().then((response) => {
                this._renderListDocument(response.value);
            });
            return;
        }

        this._getListDocument().
            then((response) => {
                console.log("Karsten2 json ", response);
                this._renderListDocument(response.value);
            });
    }

    private _renderListDataAsync(): void {
        this._getMockListDataData().then((response) => {
            this._renderDataList(response.value);
        });
        return;
    }

    private _renderList(items: ISPList[]): void {

        let html: string = "";
        items.forEach((item: ISPList) => {
            html += `<UL><LI>${item.Title} ${item.Id}</LI></ULZ`;
        });

        const listContainer: Element = this.domElement.querySelector("#spListContainer");
        listContainer.innerHTML = html;
    }

    private _renderListDocument(items: ISPListDocument[]): void {
        console.log("KarsteN");

        let html: string = "Docuements <TABLE><TR><TD>Id</TD><TD>Title</TD><TD>Deliverables</TD><TD>Phase</TD><TD>Status1</TD><TD>Link</TD></TR>";
        items.forEach((item: ISPListDocument) => {
            html += `<TR><TD>${item.Id}</TD><TD>${item.Title}</TD><TD>${item.Deliverables}</TD><TD>${item.Phase}</TD><TD>${item.Status1}</TD><TD><a href="${item.ServerRedirectedEmbedUrl}">link</a></TD></TR>`;
        });

        html += "</TABLE>";

        //let html = "Karsten<BR>" + items;

        // const listContainer: Element = this.domElement.querySelector("#spListItemContainer");
        // listContainer.innerHTML = html;
    }

    private _renderDataList(items: ISPListData[]): void {

        let html: string = "<TABLE><TR>";
        items.forEach((item: ISPListData) => {
            if (item.Status === "Approved") {
                html += `<TD class="${styles.tdgreen}">${item.Title}</TD>`;
            }
            else if (item.Status === "Draft") {
                html += `<TD class="${styles.tdorange}">${item.Title}</TD>`;
            }
            else {
                html += `<TD class="${styles.tdgrey}">${item.Title}</TD>`;
            }
        });

        html += "</TR></TABLE>";

        const listContainer: Element = this.domElement.querySelector("#spListDataContainer");
        listContainer.innerHTML = html;
    }

    private _getListData(): Promise<ISPLists> {
        /*?$filter=Hidden eq true*/
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists", SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            });
    }

    private _getMockListData(): Promise<ISPLists> {
        return MockHttpClient.get().then((data: ISPList[]) => {
            var listData: ISPLists = { value: data };
            return listData;
        }) as Promise<ISPLists>;
    }

    private _getListDocument(): Promise<ISPListDocuments> {
        /*"/_api/web/lists/getbytitle('Documents')/Items"*/
        /*let camlQueryPayLoad: any = {  
            query: {  
                __metadata: { type: “SP.CamlQuery” },  
                ViewXml: query  
            }  
        };  
     
        let spOpts = {                  
            body: JSON.stringify(camlQueryPayLoad)  
        };  */
        // return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists(guid'b28e0d0a-548b-4fbc-95f8-fac3b3b44029')/Items?$filter=Phase ne null", SPHttpClient.configurations.v1)
        // &$filter=FileRef='/sites/RossManagement/Delte dokumenter/Projekt 11 Demo' eq true
        // $select=EncodedAbsUrl,FileRef
        //return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists(guid'b28e0d0a-548b-4fbc-95f8-fac3b3b44029')/Items?$select=EncodedAbsUrl,FileRef,Id,Deliverables,Phase,Status1&$filter=startswith(FileRef, '/sites/RossManagement/Delte dokumenter/Projekt 11 Demo')", SPHttpClient.configurations.v1)
        //return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists(guid'b28e0d0a-548b-4fbc-95f8-fac3b3b44029')/Items?$filter=FileRef eq '/sites/RossManagement/Delte dokumenter/Projekt 11 Demo'", SPHttpClient.configurations.v1)
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists(guid'b28e0d0a-548b-4fbc-95f8-fac3b3b44029')/Items?$select=*,EncodedAbsUrl,FileRef,Id,Deliverables,Phase,Status1&$filter=FSObjType eq 1", SPHttpClient.configurations.v1)

            .then((response: SPHttpClientResponse) => {
                console.log("Svar" + response.status);
                return response.json();
            });
    }

    private _getMockListDocument(): Promise<ISPListDocuments> {
        return MockHttpClient.getDocuments().then((data: ISPListDocument[]) => {
            var listData: ISPListDocuments = { value: data };
            return listData;
        }) as Promise<ISPListDocuments>;
    }

    private _getMockListDataData(): Promise<ISPListDatas> {
        return MockHttpClient.getData().then((data: ISPListData[]) => {
            var listData: ISPListDatas = { value: data };
            return listData;
        }) as Promise<ISPListDatas>;
    }

    public render(): void {
        this.fetchDocumentsGuidAsync();

        this.domElement.innerHTML = `
      <div class="${ styles.rossWayWebpart}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <p class="${ styles.description}">${escape(this.context.pageContext.web.title)}</p>
              <p class="${ styles.description}">${escape(this.context.pageContext.web.absoluteUrl)}</p>
            </div>
          </div>
          <div id="spListContainer"></div>
          <div id="spListDataContainer"></div>
          <div id="spListItemContainer"></div>
        </div>
      </div>`;

        //this._renderListAsync();
        this._renderListDataAsync();
        this._renderListDocumentAsync();
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        if (!this.projectsFetched) {
            this.fetchProjects().then((response) => {
                this.projectsOptions = response;
                this.projectsFetched = true;
                // now refresh the property pane, now that the promise has been resolved..
                this.context.propertyPane.refresh();

            });
        }


        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('description', {
                                    label: strings.DescriptionFieldLabel
                                }),
                                PropertyPaneDropdown('project', {
                                    label: 'Project',
                                    options: this.projectsOptions
                                }),
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
