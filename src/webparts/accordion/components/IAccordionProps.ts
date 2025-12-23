export interface IAccordionProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
export interface IAccordionItem {
  id: string;
  title: string;
  content: string;
  isOpen: boolean;
}

export interface IAccordionWebPartProps {
  items: IAccordionItem[];
  accordionMode: 'single' | 'openAll';
}
