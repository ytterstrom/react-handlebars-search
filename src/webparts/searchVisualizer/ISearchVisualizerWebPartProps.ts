export interface ISearchVisualizerWebPartProps {
    title: string;
    query: string;
    maxResults: number;
    sorting: string;
    debug: boolean;
    external: string;
    scriptloading: boolean;
    duplicates: boolean;
    privateGroups: boolean;
    personalized: boolean;
    personalizedProperty: string;
    managedProperty: string;
}
