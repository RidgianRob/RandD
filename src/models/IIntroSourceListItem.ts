export interface IIntroSourceListItem {
  ['@odata.type']?: string;
  ['@odata.etag']?: string;
  Id: number;
  Title: string;
  introSourceJV: string,
  dealer?: string;
  introSource: string;
  reviewDate?: string;
  subArea?: string;
}