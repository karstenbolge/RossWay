import { Version } from '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneDropdown,
} from '@microsoft/sp-webpart-base';

import styles from './RossWayWebpartWebPart.module.scss';
import * as strings from 'RossWayWebpartWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';

export interface IRossWayWebpartWebPartProps {
    description: string;
    project: string;
}

interface ISPProject {
    FileRef: string;
    Id: string;
}

enum RossPhases {
    Initialization = "Initialization",
    BasisDesign = "Basis Design",
    ChallengeDesign = "Challenge Design",
    DetailledPlanning = "Detailed Planning",
    DesignOptimization = "Design Optimization",
    FinalDesign = "Final Design",
    Execution = "Execution",
    Learning = "Learning",
}

enum RossDeliverables {
    DrillingProject = "Drilling Project",
    ContinuousTasks = "Continuous Tasks",
    Approvals = "Approvals",
    Milestones = "Milestones",
    Checklists = "Checklists",
    LegislativeRequirements = "Legislative Requirements",
    OfficialApprovals = "Official Approvals",
}

enum RossStatus {
    Notstarted = "Not started",
    Approved = "Approved",
}

export default class RossWayWebpartWebPart extends BaseClientSideWebPart<IRossWayWebpartWebPartProps> {
    private projectsFetched: boolean;
    private projectsOptions: IPropertyPaneDropdownOption[];

    private documentsGuid: string;

    private fetchLists(url: string): Promise<any> {
        return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1 /*, {
            headers: {
                'Accept-Language': 'en-US,en'
            }
        }*/)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            });
    }

    private async fetchProjects(): Promise<IPropertyPaneDropdownOption[]> {
        const response = await this.fetchLists(this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists(guid'" + this.documentsGuid + "')/Items?$select=*,EncodedAbsUrl,FileRef,Id&$filter=FSObjType eq 1");
        var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        response.value.map((list: ISPProject) => {
            options.push({ key: list.FileRef.substr(list.FileRef.lastIndexOf('/') + 1), text: list.FileRef.substr(list.FileRef.lastIndexOf('/') + 1) });
        });
        return options;
    }

    private fetchDocumentsGuidAsync(): void {
        this.fetchLists(this.context.pageContext.web.absoluteUrl + "/_api/web/lists").
            then((response) => {
                this.documentsGuid = response.value.filter((item) => {
                    if (item.Title === "Documents") return true;
                })[0].Id;

                if (!this.properties.project) {
                    this.domElement.querySelector("#spRossWay").innerHTML = "<H2>No project selected.</H2>";
                    return;
                }

                this.fetchLists(this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists(guid'" + this.documentsGuid + "')/Items?$select=FSObjType,EncodedAbsUrl,FileRef,Id,RossDeliverables,RossPhase,RossStatus,RossWay&$filter=startswith(FileRef, '/sites/RossWay/Shared documents/" + this.properties.project + "/') and (RossWay eq 'Yes' or RossWay eq 'Ja' or RossWay eq 'True')")
                    .then((response2) => {
                        if (response2.error) {
                            this.domElement.querySelector("#spRossWay").innerHTML = "<i>" + response2.error.message + "</I>";
                            return;
                        }

                        let color: number[][] = [[0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0]];
                        let col: number;
                        let row: number;

                        let htmlUncategorized: string = "<BR><H2>RossWay Uncategorized Docuements</H2><TABLE><TR><TD><B>Title</B></TD><TD><B>RossDeliverables</B></TD><TD><B>RossPhase</B></TD><TD><B>RossStatus<B></TD></TR>";
                        response2.value.map((list) => {
                            // do not look at folders only files
                            if (list.FSObjType === 0) {
                                col = -1;
                                if (list.RossPhase === RossPhases.Initialization) col = 0;
                                else if (list.RossPhase === RossPhases.BasisDesign) col = 1;
                                else if (list.RossPhase === RossPhases.ChallengeDesign) col = 2;
                                else if (list.RossPhase === RossPhases.DetailledPlanning) col = 3;
                                else if (list.RossPhase === RossPhases.DesignOptimization) col = 4;
                                else if (list.RossPhase === RossPhases.FinalDesign) col = 5;
                                else if (list.RossPhase === RossPhases.Execution) col = 6;
                                else if (list.RossPhase === RossPhases.Learning) col = 7;

                                row = -1;
                                if (list.RossDeliverables === RossDeliverables.DrillingProject) row = 0;
                                else if (list.RossDeliverables === RossDeliverables.ContinuousTasks) row = 1;
                                else if (list.RossDeliverables === RossDeliverables.Approvals) row = 2;
                                else if (list.RossDeliverables === RossDeliverables.Milestones) row = 3;
                                else if (list.RossDeliverables === RossDeliverables.Checklists) row = 4;
                                else if (list.RossDeliverables === RossDeliverables.LegislativeRequirements) row = 5;
                                else if (list.RossDeliverables === RossDeliverables.OfficialApprovals) row = 6;

                                // Any status is now valid
                                if (row !== -1 && col !== -1 /*&& (list.RossStatus === RossStatus.Notstarted || list.RossStatus === RossStatus.Approved)*/) {
                                    if (color[row][col] === 0) {
                                        if (list.RossStatus === RossStatus.Notstarted) color[row][col] = 1;
                                        else if (list.RossStatus === RossStatus.Approved) color[row][col] = 2;
                                        else color[row][col] = 3;
                                    }
                                    else if (color[row][col] === 1) {
                                        if (list.RossStatus !== RossStatus.Notstarted) color[row][col] = 3;
                                    }
                                    else if (color[row][col] === 2) {
                                        if (list.RossStatus !== RossStatus.Approved) color[row][col] = 3;
                                    }
                                } else {
                                    htmlUncategorized += `<TR><TD><a href="${list.EncodedAbsUrl}">${list.FileRef.substr(list.FileRef.lastIndexOf('/') + 1)}</a></TD><TD>${list.RossDeliverables}</TD><TD>${list.RossPhase}</TD><TD>${list.RossStatus}</TD></TR>`;
                                }
                            }
                        });
                        htmlUncategorized += "</TABLE>";

                        let html: string = "";

                        let headerRow: string = "<TABLE><TR><TD><H2>" + this.properties.project + "</H2></TD>";
                        let headerColor: number[] = [0, 0, 0, 0, 0, 0, 0, 0];
                        let tableBody = "";
                        for (row = 0; row < 7; row++) {
                            tableBody += "<TR><TD class=" + styles.tdblue + ">";
                            if (row === 0) tableBody += RossDeliverables.DrillingProject;
                            if (row === 1) tableBody += RossDeliverables.ContinuousTasks;
                            if (row === 2) tableBody += RossDeliverables.Approvals;
                            if (row === 3) tableBody += RossDeliverables.Milestones;
                            if (row === 4) tableBody += RossDeliverables.Checklists;
                            if (row === 5) tableBody += RossDeliverables.LegislativeRequirements;
                            if (row === 6) tableBody += RossDeliverables.OfficialApprovals;
                            tableBody += "</TD>";
                            for (col = 0; col < 8; col++) {
                                if (row === 0) headerColor[col] = color[row][col];

                                if (color[row][col] === 0) {
                                    tableBody += "<TD class=" + styles.tdred + "> </TD>";
                                    // no document does not change header color
                                }
                                else if (color[row][col] === 1) {
                                    tableBody += "<TD class=" + styles.tdred + "> </TD>";
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
                            /* no coloring of headlines as first assumed, keeping the code if needed again 
                            if (headerColor[col] === 0) headerRow += "<TD class=" + styles.tdred + ">";
                            else if (headerColor[col] === 1) headerRow += "<TD class=" + styles.tdred + ">";
                            else if (headerColor[col] === 2) headerRow += "<TD class=" + styles.tdgreen + ">";
                            else if (headerColor[col] === 3) headerRow += "<TD class=" + styles.tdorange + ">"; */

                            headerRow += "<TD class=" + styles.tdblue + "><div style=\"width:70px\">";
                            if (col === 0) headerRow += RossPhases.Initialization.replace(" ", "<BR>");
                            else if (col === 1) headerRow += RossPhases.BasisDesign.replace(" ", "<BR>");
                            else if (col === 2) headerRow += RossPhases.ChallengeDesign.replace(" ", "<BR>");
                            else if (col === 3) headerRow += RossPhases.DetailledPlanning.replace(" ", "<BR>");
                            else if (col === 4) headerRow += RossPhases.DesignOptimization.replace(" ", "<BR>");
                            else if (col === 5) headerRow += RossPhases.FinalDesign.replace(" ", "<BR>");
                            else if (col === 6) headerRow += RossPhases.Execution.replace(" ", "<BR>");
                            else if (col === 7) headerRow += RossPhases.Learning.replace(" ", "<BR>");

                            headerRow += "</div></TD>";
                        }

                        headerRow += "</TR>";

                        html += headerRow + tableBody + "</TABLE>" + htmlUncategorized;

                        const listContainer: Element = this.domElement.querySelector("#spRossWay");
                        listContainer.innerHTML = html;
                    });
            });
    }

    public render(): void {
        this.fetchDocumentsGuidAsync();

        this.domElement.innerHTML = `<div class="${styles.rossWayWebpart}"><div id="spRossWay"></div><div id="spListContainer"></div></div>`;
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
