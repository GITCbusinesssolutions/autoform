export interface ServiceM8Field {
  label: string;
  type: 'text' | 'number' | 'date' | 'checkbox' | 'select' | 'signature' | 'photo';
  options?: string[];
  required: boolean;
  additionalDetails?: string;
}

export interface AppSettings {
  headerText: string;
  footerText: string;
  companyName: string;
  showTable: boolean;
  servicem8ApiKey?: string;
}

export interface GeneratedForm {
  title: string;
  badgeName: string;
  description: string;
  fields: ServiceM8Field[];
  docxContent: {
    sections: {
      title: string;
      content: string;
      isStandardHeader?: boolean;
      isStandardFooter?: boolean;
      layout?: 'text' | 'table' | 'grid';
    }[];
  };
}
