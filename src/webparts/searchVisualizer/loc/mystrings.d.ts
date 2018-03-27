declare interface ISearchVisualizerStrings {
    /* Fields */
    PropertyPaneDescription: string;
    QueryGroupName: string;
    TemplateGroupName: string;
    TitleFieldLabel: string;
    QueryFieldLabel: string;
    QueryFieldDescription: string;
    FieldsMaxResults: string;
    SortingFieldLabel: string;
    DebugFieldLabel: string;
    DebugFieldLabelOn: string;
    DebugFieldLabelOff: string;
    ExternalFieldLabel: string;
    ScriptloadingFieldLabel: string;
    ScriptloadingFieldLabelOn: string;
    ScriptloadingFieldLabelOff: string;
    DuplicatesFieldLabel: string;
    DuplicatesFieldLabelOn: string;
    DuplicatesFieldLabelOff: string;
    PrivateGroupsFieldLabel: string;
    PrivateGroupsFieldLabelOn: string;
    PrivateGroupsFieldLabelOff: string;
    PersonalizedFieldLabel: string;
    PersonalizedFieldLabelOn: string;
    PersonalizedFieldLabelOff: string;
    PersonalizedPropertyFieldLabel: string;
    ManagedPropertyFieldLabel: string;

    /* Validation */
    QueryValidationEmpty: string;
    TemplateValidationEmpty: string;
    TemplateValidationHTML: string;
    PersonalizedPropertyValidationEmpty:string;

    /* Dialog */
    ScriptsDialogHeader: string;
    ScriptsDialogSubText: string;
}

declare module 'searchVisualizerStrings' {
    const strings: ISearchVisualizerStrings;
    export = strings;
}
