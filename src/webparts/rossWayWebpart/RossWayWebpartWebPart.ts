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

enum Phases {
    Initialization = "Initialization",
    BasisDesign = "Basis Design",
    ChallengeDesign = "Challenge Design",
    DetailledPlanning = "Detailed Planning",
    DesignOptimization = "Design Optimization",
    FinalDesign = "Final Design",
    Execution = "Execution",
    Learning = "Learning",
}

enum Deliverables {
    BorePlanning = "Bore Planning",
    ContinuousTasks = "Continuous Tasks",
    Milestones = "Milestones",
    FormalRequirements = "Formal Requirements",
    OfficialApprovements = "Official Approvements",
}

enum Status {
    Notstarted = "Not started",
    Approved = "Approved",
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
                        let color: number[][] = [[0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0]];
                        let col: number;
                        let row: number;

                        let htmlUncategorized: string = "<BR><H2>Uncategorized Docuements</H2><TABLE><TR><TD>Title</TD><TD>Deliverables</TD><TD>Phase</TD><TD>Status1</TD></TR>";
                        response2.value.map((list) => {
                            if (list.FSObjType === 0) {
                                col = -1;
                                if (list.Phase === Phases.Initialization) col = 0;
                                else if (list.Phase === Phases.BasisDesign) col = 1;
                                else if (list.Phase === Phases.ChallengeDesign) col = 2;
                                else if (list.Phase === Phases.DetailledPlanning) col = 3;
                                else if (list.Phase === Phases.DesignOptimization) col = 4;
                                else if (list.Phase === Phases.FinalDesign) col = 5;
                                else if (list.Phase === Phases.Execution) col = 6;
                                else if (list.Phase === Phases.Learning) col = 7;

                                row = -1;
                                if (list.Deliverables === Deliverables.BorePlanning) row = 0;
                                else if (list.Deliverables === Deliverables.ContinuousTasks) row = 1;
                                else if (list.Deliverables === Deliverables.Milestones) row = 2;
                                else if (list.Deliverables === Deliverables.FormalRequirements) row = 3;
                                else if (list.Deliverables === Deliverables.OfficialApprovements) row = 4;

                                if (row !== -1 && col !== -1 && (list.Status1 === Status.Notstarted || list.Status1 === Status.Approved)) {
                                    if (color[row][col] === 0) {
                                        if (list.Status1 === Status.Notstarted) color[row][col] = 1;
                                        else if (list.Status1 === Status.Approved) color[row][col] = 2;
                                    }
                                    else if (color[row][col] === 1) {
                                        if (list.Status1 === Status.Approved) color[row][col] = 3;
                                    }
                                    else if (color[row][col] === 2) {
                                        if (list.Status1 === Status.Notstarted) color[row][col] = 3;
                                    }
                                } else {
                                    htmlUncategorized += `<TR><TD><a href="${list.EncodedAbsUrl}">${list.FileRef.substr(list.FileRef.lastIndexOf('/') + 1)}</a></TD><TD>${list.Deliverables}</TD><TD>${list.Phase}</TD><TD>${list.Status1}</TD></TR>`;
                                }
                            }
                        });
                        htmlUncategorized += "</TABLE>";

                        let html: string = "";

                        let headerRow: string = "<TABLE><TR><TD><H2>" + this.properties.project + "</H2></TD>";
                        let headerColor: number[] = [0, 0, 0, 0, 0, 0, 0, 0];
                        let tableBody = "";
                        for (row = 0; row < 5; row++) {
                            tableBody += "<TR><TD class=" + styles.tdblue + ">";
                            if (row === 0) tableBody += Deliverables.BorePlanning;
                            if (row === 1) tableBody += Deliverables.ContinuousTasks;
                            if (row === 2) tableBody += Deliverables.Milestones;
                            if (row === 3) tableBody += Deliverables.FormalRequirements;
                            if (row === 4) tableBody += Deliverables.OfficialApprovements;
                            tableBody += "</TD>";
                            for (col = 0; col < 8; col++) {
                                if (row === 0) headerColor[col] = color[row][col];

                                if (color[row][col] === 0) {
                                    tableBody += "<TD class=" + styles.tdgrey + "> </TD>";
                                    // no document does not change header color
                                }
                                else if (color[row][col] === 1) {
                                    tableBody += "<TD class=" + styles.tdgrey + "> </TD>";
                                    if (headerColor[col] === 0) headerColor[col] = 1;
                                    else if (headerColor[col] == 2) headerColor[col] = 3;
                                }
                                else if (color[row][col] === 2) {
                                    tableBody += "<TD class=" + styles.tdgreen + "> </TD>";
                                    if (headerColor[col] === 0) headerColor[col] = 2;
                                    else if (headerColor[col] == 1) headerColor[col] = 3;
                                }
                                else if (color[row][col] === 3) {
                                    tableBody += "<TD class=" + styles.tdorange + "> </TD>";
                                    headerColor[col] = 3;
                                }

                            }
                            tableBody += "</TR>";
                        }

                        for (col = 0; col < 8; col++) {
                            if (headerColor[col] === 0) headerRow += "<TD class=" + styles.tdgrey + ">";
                            else if (headerColor[col] === 1) headerRow += "<TD class=" + styles.tdgrey + ">";
                            else if (headerColor[col] === 2) headerRow += "<TD class=" + styles.tdgreen + ">";
                            else if (headerColor[col] === 3) headerRow += "<TD class=" + styles.tdorange + ">";

                            if (col === 0) headerRow += Phases.Initialization;
                            else if (col === 1) headerRow += Phases.BasisDesign;
                            else if (col === 2) headerRow += Phases.ChallengeDesign;
                            else if (col === 3) headerRow += Phases.DetailledPlanning;
                            else if (col === 4) headerRow += Phases.DesignOptimization;
                            else if (col === 5) headerRow += Phases.FinalDesign;
                            else if (col === 6) headerRow += Phases.Execution;
                            else if (col === 7) headerRow += Phases.Learning;

                            headerRow += "</TD>";
                        }

                        headerRow += "</TR>";

                        html += headerRow + tableBody + "</TABLE>" + htmlUncategorized;


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
        let color: number[][];
        let col: number;
        let row: number;

        let html: string = "Docuements <TABLE><TR><TD>Id</TD><TD>Title</TD><TD>Deliverables</TD><TD>Phase</TD><TD>Status1</TD><TD>Link</TD><TD>col</TD><TD>row</TD></TR>";
        items.forEach((item: ISPListDocument) => {
            col = -1;
            if (item.Phase === Phases.Initialization) col = 0;
            if (item.Phase === Phases.BasisDesign) col = 1;
            if (item.Phase === Phases.ChallengeDesign) col = 2;
            if (item.Phase === Phases.DetailledPlanning) col = 3;
            if (item.Phase === Phases.DesignOptimization) col = 4;
            if (item.Phase === Phases.FinalDesign) col = 5;
            if (item.Phase === Phases.Execution) col = 6;
            if (item.Phase === Phases.Learning) col = 7;

            row = -1;
            if (item.Phase === Deliverables.BorePlanning) row = 0;
            if (item.Phase === Deliverables.ContinuousTasks) row = 1;
            if (item.Phase === Deliverables.Milestones) row = 2;
            if (item.Phase === Deliverables.FormalRequirements) row = 3;
            if (item.Phase === Deliverables.OfficialApprovements) row = 4;

            html += `<TR><TD>${item.Id}</TD><TD>${item.Title}</TD><TD>${item.Deliverables}</TD><TD>${item.Phase}</TD><TD>${item.Status1}</TD><TD><a href="${item.ServerRedirectedEmbedUrl}">link</a></TD><TD>${col}</TD><TD>${row}</TD></TR>`;
        });

        html += "</TABLE>";

        //let html = "Karsten<BR>" + items;

        // const listContainer: Element = this.domElement.querySelector("#spListItemContainer");
        // listContainer.innerHTML = html;
    }

    private _renderDataList(items: ISPListData[]): void {

        let html: string = "KARSTEN !!!<TABLE><TR>";
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

        const unused = `
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

        this.domElement.innerHTML = `<div class="${styles.rossWayWebpart}"><div id="spListItemContainer"></div></div>`;
        //this._renderListAsync();
        //this._renderListDataAsync();
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
