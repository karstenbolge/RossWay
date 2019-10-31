import { ISPList, ISPListData, ISPListDocument } from "./RossWayWebpartWebPart";

export default class MockHttpClient {
    private static _items: ISPList[] = [{ Title: "Mock list 1", Id: "1" },
    { Title: "Mock 2", Id: "2" },
    { Title: "Mock 3", Id: "3" }];

    public static get(): Promise<ISPList[]> {
        return new Promise<ISPList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }

    private static _dataItems: ISPListData[] = [{ Title: "Initialisation", Status: "Approved", Id: "1" },
    { Title: "Basis Design", Status: "Draft", Id: "2" },
    { Title: "Challenge Design", Status: "Not started", Id: "3" },
    { Title: "Detailled Planning", Status: "Not started", Id: "4" },
    { Title: "Design Optimization", Status: "Not started", Id: "5" },
    { Title: "Final Design", Status: "Not started", Id: "6" },
    { Title: "Execution", Status: "Not started", Id: "7" },
    { Title: "Learning", Status: "Not started", Id: "8" },
    ];

    public static getData(): Promise<ISPListData[]> {
        return new Promise<ISPListData[]>((resolve) => {
            resolve(MockHttpClient._dataItems);
        });
    }

    private static _documents: ISPListDocument[] = [{
        Title: "Doc I", Id: "1", Deliverables: "YES", Phase: "DEl", Status1: "fino", ServerRedirectedEmbedUrl: "string",

    }];

    public static getDocuments(): Promise<ISPListDocument[]> {
        return new Promise<ISPListDocument[]>((resolve) => {
            resolve(MockHttpClient._documents);
        });
    }

}