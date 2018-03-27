import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import pnp from "sp-pnp-js";
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneToggle,
    PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import * as strings from 'searchVisualizerStrings';
import SearchVisualizer from './components/SearchVisualizer';
import { ISearchVisualizerProps } from './components/ISearchVisualizerProps';
import { ISearchVisualizerWebPartProps } from './ISearchVisualizerWebPartProps';
import { SPComponentLoader } from '@microsoft/sp-loader';

export default class SearchVisualizerWebPart extends BaseClientSideWebPart<ISearchVisualizerWebPartProps> {
    private personalizedPropertyDisabled: boolean = true;
    private managedPropertyDisabled: boolean = true;
    constructor() {
        super();

        // Load the core UI Fabric styles
        SPComponentLoader.loadCss('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/6.0.0/css/fabric-6.0.0.scoped.min.css');
      }

    public render(): void {
        const element: React.ReactElement<ISearchVisualizerProps> = React.createElement(
            SearchVisualizer,
            {
                title: this.properties.title,
                query: this.properties.query,
                maxResults: this.properties.maxResults,
                sorting: this.properties.sorting,
                debug: this.properties.debug,
                external: this.properties.external,
                scriptloading: this.properties.scriptloading,
                duplicates: this.properties.duplicates,
                privateGroups: this.properties.privateGroups,
                personalized: this.properties.personalized,
                personalizedProperty: this.properties.personalizedProperty,
                managedProperty: this.properties.managedProperty,
                context: this.context
            }
        );

        ReactDom.render(element, this.domElement);
    }
    protected onInit(): Promise<void> {

        return super.onInit().then(_ => {
          pnp.setup({
            spfxContext: this.context
          });
        });
      }
    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.QueryGroupName,
                            groupFields: [
                                PropertyPaneTextField('query', {
                                    label: strings.QueryFieldLabel,
                                    description: strings.QueryFieldDescription,
                                    multiline: true,
                                    onGetErrorMessage: this._queryValidation,
                                    deferredValidationTime: 500
                                }),
                                PropertyPaneSlider('maxResults', {
                                    label: strings.FieldsMaxResults,
                                    min: 1,
                                    max: 50
                                }),
                                PropertyPaneTextField('sorting', {
                                    label: strings.SortingFieldLabel
                                }),
                                PropertyPaneToggle('duplicates', {
                                    label: strings.DuplicatesFieldLabel,
                                    onText: strings.DuplicatesFieldLabelOn,
                                    offText: strings.DuplicatesFieldLabelOff
                                }),
                                PropertyPaneToggle('privateGroups', {
                                    label: strings.PrivateGroupsFieldLabel,
                                    onText: strings.PrivateGroupsFieldLabelOn,
                                    offText: strings.PrivateGroupsFieldLabelOff
                                }),
                                PropertyPaneToggle('personalized', {
                                    label: strings.PersonalizedFieldLabel,
                                    onText: strings.PersonalizedFieldLabelOn,
                                    offText: strings.PersonalizedFieldLabelOff
                                }),
                                PropertyPaneTextField('personalizedProperty', {
                                    label: strings.PersonalizedPropertyFieldLabel,
                                    disabled: this.personalizedPropertyDisabled

                                }),
                                PropertyPaneTextField('managedProperty', {
                                    label: strings.ManagedPropertyFieldLabel,
                                    disabled: this.managedPropertyDisabled

                                })
                            ],
                            isCollapsed: true
                        },
                        {
                            groupName: strings.TemplateGroupName,
                            groupFields: [
                                PropertyPaneTextField('title', {
                                    label: strings.TitleFieldLabel
                                }),
                                PropertyPaneToggle('debug', {
                                    label: strings.DebugFieldLabel,
                                    onText: strings.DebugFieldLabelOn,
                                    offText: strings.DebugFieldLabelOff
                                }),
                                PropertyPaneTextField('external', {
                                    label: strings.ExternalFieldLabel,
                                    onGetErrorMessage: this._externalTemplateValidation.bind(this)
                                }),
                                PropertyPaneToggle('scriptloading', {
                                    label: strings.ScriptloadingFieldLabel,
                                    onText: strings.ScriptloadingFieldLabelOn,
                                    offText: strings.ScriptloadingFieldLabelOff
                                })
                            ],
                            isCollapsed: true
                        }
                    ],
                    displayGroupsAsAccordion: true
                }
            ]
        };
    }

    /**
     * Validating the query property
     *
     * @param value
     */
    private _queryValidation(value: string): string {
        // Check if a URL is specified
        if (value.trim() === "") {
            return strings.QueryValidationEmpty;
        }

        return '';
    }
    private _personalizedPropertyValidation(value: string):string {
        debugger;
        if (value.trim() === "" && this.properties.personalized ===true) {
            return strings.PersonalizedPropertyValidationEmpty;
        }
    }

    /**
     * Validating the external template property
     *
     * @param value
     */
    private _externalTemplateValidation(value: string): string {
        // If debug template is set to off, user needs to specify a template URL
        if (!this.properties.debug) {
            // Check if a URL is specified
            if (value.trim() === "") {
                return strings.TemplateValidationEmpty;
            }
            // Check if a HTML file is referenced
            if (value.toLowerCase().indexOf('.html') === -1) {
                return strings.TemplateValidationHTML;
            }
        }

        return '';
    }

    /**
	 * Prevent from changing the query on typing
	 */
    protected get disableReactivePropertyChanges(): boolean {
        return true;
    }
    protected onPropertyPaneConfigurationStart(): void {

        this.personalizedPropertyDisabled = !this.properties.personalized;
        this.managedPropertyDisabled = !this.properties.personalized;
    }
    protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {

        if(propertyPath==="personalized"  ){

            this.personalizedPropertyDisabled =!newValue;
            this.managedPropertyDisabled = !newValue;
        }
    }
}
