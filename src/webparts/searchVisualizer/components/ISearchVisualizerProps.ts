import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISearchVisualizerProps {
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
    context: WebPartContext;
}

export interface ISearchVisualizerState {
    loading?: boolean;
    template?: string;
    error?: string;
    showError?: boolean;
    showScriptDialog?: boolean;
}

export interface IMetadata {
    fields: string[];
}

export interface ISPUser {
    username?: string;
    displayName?: string;
    email?: string;
}

export interface ISPUrl {
    url?: string;
    description?: string;
}


