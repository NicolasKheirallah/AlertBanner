export interface IFormErrors {
  title?: string;
  description?: string;
  AlertType?: string;
  linkUrl?: string;
  linkDescription?: string;
  targetSites?: string;
  scheduledStart?: string;
  scheduledEnd?: string;
  [key: string]: string | undefined;
}
