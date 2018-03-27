import * as moment from 'moment';
import * as Handlebars from 'handlebars';
import { ISPUser, ISPUrl } from './../components/ISearchVisualizerProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import pnp from "sp-pnp-js";

export default class HBSharePointHelpers {
    constructor(private _context: WebPartContext) {
        Handlebars.registerHelper('splitDisplayNames', this._splitDisplayNames);
        Handlebars.registerHelper('splitSPUser', this._splitSPUser);
        Handlebars.registerHelper('splitSPTaxonomy', this._splitSPTaxonomy);
        Handlebars.registerHelper('splitSPUrl', this._splitSPUrl);
        Handlebars.registerHelper('formatDate', this._formatDate);
        Handlebars.registerHelper('returnday', this._returnday);
        Handlebars.registerHelper('returnMonthName', this._returnMonthName);
        Handlebars.registerHelper('evenRow', this._evenRow);
        Handlebars.registerHelper('currentWebUrl', this._currentWebUrl);
    }

    /**
     * Initialize the class
     * @param _context
     */
    public static init(_context: WebPartContext) {
        const instance = new HBSharePointHelpers(_context);
    }

    /**
     * SharePoint helper to split the displaynames of for example the Author field (user1;user2...)
     * @param displayNames
     */
    private _splitDisplayNames = (displayNames) => {
        if (displayNames == null && displayNames.indexOf(';') == -1) {
            return null;
        }

        return displayNames.split(';').join(", ");
    }

    /**
     * SharePoint helper to split SPUserField (?multiple) into a string.
     * The template provide the property which will be returned.
     * @param userFieldValue
     * @param propertyRequested
     */
    private _splitSPUser = (userFieldValue, propertyRequested) => {
        if (userFieldValue == null)
            return null;

        const retValue: string[] = [];
        let userFieldValueArray = userFieldValue.split(';').forEach(user => {
            let userValues = user.split(' | ');
            let spuser: ISPUser = {
                displayName: userValues[1],
                email: userValues[0]
            };
            retValue.push(spuser[propertyRequested]);
        });

        return retValue.join(', ');
    }

    /**
     * SharePoint helper to split the taxonomy name
     * @param taxonomyFieldValue
     */
    private _splitSPTaxonomy = (taxonomyFieldValue) => {
        if (taxonomyFieldValue == null)
            return null;

        const retValue: string[] = [];

        let taxonomyFieldValueArray = taxonomyFieldValue.split(';').forEach(taxonomy => {
            if (taxonomy.indexOf('L0|') !== -1) {
                retValue.push(taxonomy.split('|').pop());
            }
        });
        return retValue.join(', ');
    }

    /**
     * SharePoint helper to split url/desciption
     * @param urlFieldValue
     * @param propertyRequested
     */
    private _splitSPUrl = (urlFieldValue, propertyRequested) => {
        if (urlFieldValue == null)
            return null;

        let spurl: ISPUrl = {
            url: urlFieldValue.split(',')[0],
            description: urlFieldValue.split(',')[1]
        };
        return spurl[propertyRequested];
    }
    private _formatDate = (date, format) => {
        var offset = moment().utcOffset();

        return moment.utc(date).utcOffset(offset).format(format);

    }
    private _returnMonthName = (date) => {

        var parsedate = new Date(date);

        var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ];
        var month = monthNames[parsedate.getMonth()];

        return month;
    }
    private _returnday = (date) => {

        var parsedate = new Date(date);

        var day = parsedate.getDate();

        return day;
    }
    private _evenRow = (rowNumber) => {
        if (rowNumber % 2 == 0) { return '20px;'; }
        else { return '0;'; }
    }
    private _currentWebUrl = (path) => {
        let res = window.location.href.substring(0,window.location.href.indexOf("/SitePages/"));

        return res;
    }

}
