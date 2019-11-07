import { Version } from '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneDropdown,
} from '@microsoft/sp-webpart-base';

import styles from './RossWayWebpartWebPart.module.scss';
import * as strings from 'RossWayWebpartWebPartStrings';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';

export interface IRossWayWebpartWebPartProps {
    description: string;
    project: string;
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

    private documentsGuid: string;

    private fetchLists(url: string): Promise<any> {
        return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            });
    }

    private async fetchProjects(): Promise<IPropertyPaneDropdownOption[]> {
        const response = await this.fetchLists(this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists(guid'" + this.documentsGuid + "')/Items?$select=*,EncodedAbsUrl,FileRef,Id,Deliverables,Phase,Status1&$filter=FSObjType eq 1");
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

                this.fetchLists(this.context.pageContext.web.absoluteUrl + "/_api/Web/Lists(guid'" + this.documentsGuid + "')/Items?$select=FSObjType,EncodedAbsUrl,FileRef,Id,Deliverables,Phase,Status1,RossWay&$filter=startswith(FileRef, '/sites/RossManagement/Delte dokumenter/" + this.properties.project + "/')")// /Projekt 11 Demo')")
                    .then((response2) => {
                        let color: number[][] = [[0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0]];
                        let col: number;
                        let row: number;

                        let htmlUncategorized: string = "<BR><H2>RossWay Uncategorized Docuements</H2><TABLE><TR><TD><B>Title</B></TD><TD><B>Deliverables</B></TD><TD><B>Phase</B></TD><TD><B>Status1<B></TD></TR>";
                        response2.value.map((list) => {
                            if (list.RossWay === "Yes") {
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
                            }
                        });
                        htmlUncategorized += "</TABLE>";

                        let html: string = "";

                        let headerRow: string = "<TABLE><TR><TD><H2>" + (this.properties.project || "None selected") + "</H2></TD>";
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

    public render(): void {
        this.fetchDocumentsGuidAsync();

        this.domElement.innerHTML = `<div class="${styles.rossWayWebpart}"><div id="spListItemContainer"></div><div id="spListContainer"></div></div>`;
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
