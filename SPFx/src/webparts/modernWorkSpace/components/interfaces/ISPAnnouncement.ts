export interface ISPAnnouncement{
    Id: string;
    id?: string;
    Title: string;
    Description: string;
    AuthorId?: string;
    Created: string;
    AuthorLookupId?: string;
    Author?: {Title: string};
    AttachmentFiles?: {ServerRelativeUrl: string}[];
}

export interface ISPAnnouncementItems{
    '@odata.context': string;
    value: ISPAnnouncement[];
}

export class SPAnnouncement{
    public Id: string;
    public Title: string;
    public Description: string;
    public AuthorId: string;
    public AuthorTitle: string;
    public AttachmentServerURL: string;
    public Created: Date;

    constructor(item: ISPAnnouncement){
        this.Id = item.Id ? item.Id : item.id;
        this.Title = item.Title;
        this.Description = item.Description;
        this.AuthorId = item.AuthorId ? item.AuthorId : item.AuthorLookupId;
        this.AuthorTitle = item.Author ? item.Author.Title : "";
        this.AttachmentServerURL = item.AttachmentFiles[0] ? item.AttachmentFiles[0].ServerRelativeUrl : "";
        this.Created = new Date(item.Created);
    }
}