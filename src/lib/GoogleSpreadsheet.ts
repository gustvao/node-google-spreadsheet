import _ from 'lodash';
import Axios, { AxiosError, AxiosInstance, AxiosRequestConfig, AxiosResponse } from 'axios';

// optional peer dependency - this ts-ignore should help when user is not using it
// @ts-ignore
import type { JWT, OAuth2Client } from 'google-auth-library';


import GoogleSpreadsheetWorksheet from './GoogleSpreadsheetWorksheet';
import { axiosParamsSerializer, getFieldMask } from './utils';
import { DataFilter, GridRange, NamedRangeId, RecursivePartial, SpreadsheetProperties, WorksheetGridProperties, WorksheetId, WorksheetProperties } from './shared-types';
import { PermissionRoles, PermissionsList, PublicPermissionRoles } from './drive-types';

const SHEETS_API_BASE_URL = 'https://sheets.googleapis.com/v4/spreadsheets';
const DRIVE_API_BASE_URL = 'https://www.googleapis.com/drive/v3/files';

// TODO: check JWT scopes against this list?
const GOOGLE_AUTH_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',

  // the list from the sheets v4 auth for spreadsheets.get
  // 'https://www.googleapis.com/auth/drive',
  // 'https://www.googleapis.com/auth/drive.readonly',
  // 'https://www.googleapis.com/auth/drive.file',
  // 'https://www.googleapis.com/auth/spreadsheets',
  // 'https://www.googleapis.com/auth/spreadsheets.readonly',
];


const EXPORT_CONFIG: Record<string, { singleWorksheet?: boolean }> = {
  html: {},
  zip: {},
  xlsx: {},
  ods: {},
  csv: { singleWorksheet: true },
  tsv: { singleWorksheet: true },
  pdf: { singleWorksheet: true },
}
type ExportFileTypes = keyof typeof EXPORT_CONFIG;



/** single type to handle all valid auth types */
export type GoogleSpreadsheetAuth = JWT | OAuth2Client | { apiKey: string } | { token: string };
enum AUTH_MODES {
  API_KEY,
  RAW_ACCESS_TOKEN,
  JWT,
  OAUTH,
}
function getAuthMode(auth: GoogleSpreadsheetAuth) {
  if ('apiKey' in auth) return AUTH_MODES.API_KEY;
  if ('authorize' in auth) return AUTH_MODES.JWT;
  if ('getAccessToken' in auth) return AUTH_MODES.OAUTH;
  if ('token' in auth) return AUTH_MODES.RAW_ACCESS_TOKEN;
  throw new Error('Invalid auth');
}

async function getRequestAuthConfig(auth: GoogleSpreadsheetAuth) {
  // API key only access adds passes through the key as a query param
  if ('apiKey' in auth) {
    return { params: { key: auth.apiKey }}
  }

  // all other methods pass through a bearer token in the Authorization header

  let authToken: string | null | undefined;

  // JWT
  if ('authorize' in auth) {
    await auth.authorize();
    authToken = auth.credentials.access_token;
  // OAUTH
  } else if ('getAccessToken' in auth) {
    const credentials = await auth.getAccessToken();
    authToken = credentials.token;
  } else if ('token' in auth) {
    authToken = auth.token;
  }
  if (!authToken) {
    throw new Error('Invalid auth');
  }
  return { headers: { Authorization: `Bearer ${authToken}` }};
}

export default class GoogleSpreadsheet {
  
  readonly spreadsheetId: string;
  
  public auth: GoogleSpreadsheetAuth;
  get authMode() {
    return getAuthMode(this.auth);
  }

  private _rawSheets: any;
  private _rawProperties = null as SpreadsheetProperties | null;
  private _spreadsheetUrl = null as string | null;
  private _deleted = false;

  readonly sheetsApi: AxiosInstance;
  readonly driveApi: AxiosInstance;

  constructor(spreadsheetId: string, auth: GoogleSpreadsheetAuth) {
    this.spreadsheetId = spreadsheetId;
    this.auth = auth;

    this._rawSheets = {};
    this._spreadsheetUrl = null;

    // create an axios instance with sheet root URL and interceptors to handle auth
    this.sheetsApi = Axios.create({
      baseURL: `${SHEETS_API_BASE_URL}/${spreadsheetId}`,
      paramsSerializer: axiosParamsSerializer,
    });
    this.driveApi = Axios.create({
      baseURL: `${DRIVE_API_BASE_URL}/${spreadsheetId}`,
      paramsSerializer: axiosParamsSerializer,
    })
    // have to use bind here or the functions dont have access to `this` :(
    this.sheetsApi.interceptors.request.use(this._setAxiosRequestAuth.bind(this));
    this.sheetsApi.interceptors.response.use(
      this._handleAxiosResponse.bind(this),
      this._handleAxiosErrors.bind(this)
    );
    this.driveApi.interceptors.request.use(this._setAxiosRequestAuth.bind(this));
    this.driveApi.interceptors.response.use(
      this._handleAxiosResponse.bind(this),
      this._handleAxiosErrors.bind(this)
    );

    return this;
  }


  // AUTH RELATED FUNCTIONS ////////////////////////////////////////////////////////////////////////

  // INTERNAL UTILITY FUNCTIONS ////////////////////////////////////////////////////////////////////
  async _setAxiosRequestAuth(config: AxiosRequestConfig) {
    const authConfig = await getRequestAuthConfig(this.auth);
    config.headers = {...config.headers, ...authConfig.headers };
    config.params = {...config.params, ...authConfig.params };
    return config;
  }

  async _handleAxiosResponse(response: AxiosResponse) { return response; }
  async _handleAxiosErrors(error: AxiosError) {
    // console.log(error);
    if (error.response && error.response.data) {
      // usually the error has a code and message, but occasionally not
      if (!error.response.data.error) throw error;

      const { code, message } = error.response.data.error;
      error.message = `Google API error - [${code}] ${message}`;
      throw error;
    }

    if (_.get(error, 'response.status') === 403) {
      if ('apiKey' in this.auth) {
        throw new Error('Sheet is private. Use authentication or make public. (see https://github.com/theoephraim/node-google-spreadsheet#a-note-on-authentication for details)');
      }
    }
    throw error;
  }

  async _makeSingleUpdateRequest(requestType: string, requestParams: any) {
    const response = await this.sheetsApi.post(':batchUpdate', {
      requests: [{ [requestType]: requestParams }],
      includeSpreadsheetInResponse: true,
      // responseRanges: [string]
      // responseIncludeGridData: true
    });

    this._updateRawProperties(response.data.updatedSpreadsheet.properties);
    _.each(response.data.updatedSpreadsheet.sheets, (s) => this._updateOrCreateSheet(s));
    // console.log('API RESPONSE', response.data.replies[0][requestType]);
    return response.data.replies[0][requestType];
  }

  // TODO: review these types
  // currently only used in batching cell updates
  async _makeBatchUpdateRequest(requests: any[], responseRanges?: string | string[]) {
    // this is used for updating batches of cells
    const response = await this.sheetsApi.post(':batchUpdate', {
      requests,
      includeSpreadsheetInResponse: true,
      ...responseRanges && {
        responseIncludeGridData: true,
        ...responseRanges !== '*' && { responseRanges },
      },
    });

    this._updateRawProperties(response.data.updatedSpreadsheet.properties);
    _.each(response.data.updatedSpreadsheet.sheets, (s) => this._updateOrCreateSheet(s));
  }

  _ensureInfoLoaded() {
    if (!this._rawProperties) throw new Error('You must call `doc.loadInfo()` before accessing this property');
  }

  _updateRawProperties(newProperties: SpreadsheetProperties) { this._rawProperties = newProperties; }

  _updateOrCreateSheet(sheetInfo: { properties: WorksheetProperties, data: any }) {
    const { properties, data } = sheetInfo; 
    const { sheetId } = properties;
    if (!this._rawSheets[sheetId]) {
      this._rawSheets[sheetId] = new GoogleSpreadsheetWorksheet(this, properties, data);
    } else {
      this._rawSheets[sheetId].updateRawData(properties, data);
    }
  }

  // BASIC PROPS //////////////////////////////////////////////////////////////////////////////
  _getProp(param: keyof SpreadsheetProperties) {
    this._ensureInfoLoaded();
    // ideally ensureInfoLoaded would assert that _rawProperties is in fact loaded
    // but this is not currently possible in TS - see https://github.com/microsoft/TypeScript/issues/49709
    return this._rawProperties![param];
  }
  _setProp(param: keyof SpreadsheetProperties, newVal: any) { // eslint-disable-line no-unused-vars
    throw new Error('Do not update directly - use `updateProperties()`');
  }

  get title(): SpreadsheetProperties['title'] { return this._getProp('title'); }
  get locale(): SpreadsheetProperties['locale'] { return this._getProp('locale'); }
  get timeZone(): SpreadsheetProperties['timeZone'] { return this._getProp('timeZone'); }
  get autoRecalc(): SpreadsheetProperties['autoRecalc'] { return this._getProp('autoRecalc'); }
  get defaultFormat(): SpreadsheetProperties['defaultFormat'] { return this._getProp('defaultFormat'); }
  get spreadsheetTheme(): SpreadsheetProperties['spreadsheetTheme'] { return this._getProp('spreadsheetTheme'); }
  get iterativeCalculationSettings(): SpreadsheetProperties['iterativeCalculationSettings'] { return this._getProp('iterativeCalculationSettings'); }

  set title(newVal) { this._setProp('title', newVal); }
  set locale(newVal) { this._setProp('locale', newVal); }
  set timeZone(newVal) { this._setProp('timeZone', newVal); }
  set autoRecalc(newVal) { this._setProp('autoRecalc', newVal); }
  set defaultFormat(newVal) { this._setProp('defaultFormat', newVal); }
  set spreadsheetTheme(newVal) { this._setProp('spreadsheetTheme', newVal); }
  set iterativeCalculationSettings(newVal) { this._setProp('iterativeCalculationSettings', newVal); }

  /**
   * update spreadsheet properties
   * @see https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#SpreadsheetProperties
   * */
  async updateProperties(properties: Partial<SpreadsheetProperties>) {
    await this._makeSingleUpdateRequest('updateSpreadsheetProperties', {
      properties,
      fields: getFieldMask(properties),
    });
  }

  // BASIC INFO ////////////////////////////////////////////////////////////////////////////////////
  async loadInfo(includeCells = false) {
    const response = await this.sheetsApi.get('/', {
      params: {
        ...includeCells && { includeGridData: true },
      },
    });
    this._spreadsheetUrl = response.data.spreadsheetUrl;
    this._rawProperties = response.data.properties;
    _.each(response.data.sheets, (s) => this._updateOrCreateSheet(s));
  }
  async getInfo() { return this.loadInfo(); } // alias to mimic old version

  resetLocalCache() {
    this._rawProperties = null;
    this._rawSheets = {};
  }

  // WORKSHEETS ////////////////////////////////////////////////////////////////////////////////////
  get sheetCount() {
    this._ensureInfoLoaded();
    return _.values(this._rawSheets).length;
  }

  get sheetsById(): Record<WorksheetId, GoogleSpreadsheetWorksheet> {
    this._ensureInfoLoaded();
    return this._rawSheets;
  }

  get sheetsByIndex(): GoogleSpreadsheetWorksheet[] {
    this._ensureInfoLoaded();
    return _.sortBy(this._rawSheets, 'index');
  }

  get sheetsByTitle(): Record<string, GoogleSpreadsheetWorksheet> {
    this._ensureInfoLoaded();
    return _.keyBy(this._rawSheets, 'title');
  }

  /**
   * Add new worksheet to document
   * @see https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AddSheetRequest
   * */
  async addSheet(
    properties: Partial<
      RecursivePartial<WorksheetProperties>
      & {
        headerValues: string[],
        headerRowIndex: number
      }
    > = {}
  ) {
    const response = await this._makeSingleUpdateRequest('addSheet', {
      properties: _.omit(properties, 'headerValues', 'headerRowIndex'),
    });
    // _makeSingleUpdateRequest already adds the sheet
    const newSheetId = response.properties.sheetId;
    const newSheet = this.sheetsById[newSheetId];

    if (properties.headerValues) {
      await newSheet.setHeaderRow(properties.headerValues, properties.headerRowIndex);
    }

    return newSheet;
  }

  /**
   * delete a worksheet
   * @see https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteSheetRequest
   * */
  async deleteSheet(sheetId: WorksheetId) {
    await this._makeSingleUpdateRequest('deleteSheet', { sheetId });
    delete this._rawSheets[sheetId];
  }

  // NAMED RANGES //////////////////////////////////////////////////////////////////////////////////

  /**
   * create a new named range
   * @see https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#AddNamedRangeRequest
   */
  async addNamedRange(
    /** name of new named range */
    name: string,
    /** GridRange object describing range */
    range: GridRange,
    /** id for named range (optional) */
    namedRangeId?: string
  ) {
    // TODO: add named range to local cache
    return this._makeSingleUpdateRequest('addNamedRange', {
      name,
      namedRangeId,
      range,
    });
  }

  /**
   * delete a named range
   * @see https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#DeleteNamedRangeRequest
   * */
  async deleteNamedRange(
    /** id of named range to delete */
    namedRangeId: NamedRangeId
  ) {
    // TODO: remove named range from local cache
    return this._makeSingleUpdateRequest('deleteNamedRange', { namedRangeId });
  }

  // LOADING CELLS /////////////////////////////////////////////////////////////////////////////////
  
  /** fetch cell data into local cache */
  async loadCells(
    /**
     * single filter or array of filters
     * strings are treated as A1 ranges, objects are treated as GridRange objects
     * pass nothing to fetch all cells
     * */
    filters?: DataFilter | DataFilter[]
  ) {
    // TODO: make it support DeveloperMetadataLookup objects

    

    // TODO: switch to this mode if using a read-only auth token?
    const readOnlyMode = this.authMode === AUTH_MODES.API_KEY;

    const filtersArray = _.isArray(filters) ? filters : [filters];
    const dataFilters = _.map(filtersArray, (filter) => {
      if (_.isString(filter)) {
        return readOnlyMode ? filter : { a1Range: filter };
      }
      if (_.isObject(filter)) {
        if (readOnlyMode) {
          throw new Error('Only A1 ranges are supported when fetching cells with read-only access (using only an API key)');
        }
        // TODO: make this support Developer Metadata filters
        return { gridRange: filter };
      }
      throw new Error('Each filter must be an A1 range string or a gridrange object');
    });

    let result;
    // when using an API key only, we must use the regular get endpoint
    // because :getByDataFilter requires higher access
    if (this.authMode === AUTH_MODES.API_KEY) {
      result = await this.sheetsApi.get('/', {
        params: {
          includeGridData: true,
          ranges: dataFilters,
        },
      });
    // otherwise we use the getByDataFilter endpoint because it is more flexible
    } else {
      result = await this.sheetsApi.post(':getByDataFilter', {
        includeGridData: true,
        dataFilters,
      });
    }

    const { sheets } = result.data;
    _.each(sheets, (sheet) => { this._updateOrCreateSheet(sheet); });
  }

  // EXPORTING /////////////////////////////////////////////////////////////

  /**
   * export/download helper, not meant to be called directly (use downloadAsX methods on spreadsheet and worksheet instead)
   * @internal
   */
  async _downloadAs(
    fileType: ExportFileTypes,
    worksheetId?: WorksheetId,
    returnStreamInsteadOfBuffer?: boolean
  ) {
    // see https://stackoverflow.com/questions/11619805/using-the-google-drive-api-to-download-a-spreadsheet-in-csv-format/51235960#51235960

    if (!EXPORT_CONFIG[fileType]) throw new Error(`unsupported export fileType - ${fileType}`);
    if (EXPORT_CONFIG[fileType].singleWorksheet) {
      if (worksheetId === undefined) throw new Error(`Must specify worksheetId when exporting as ${fileType}`);
    } else {
      if (worksheetId) throw new Error(`Cannot specify worksheetId when exporting as ${fileType}`);
    }

    // google UI shows "html" but passes through "zip"
    if (fileType === 'html') fileType = 'zip';

    if (!this._spreadsheetUrl) throw new Error('Cannot export sheet that is not fully loaded');

    const exportUrl = this._spreadsheetUrl.replace('/edit', '/export');
    const response = await this.sheetsApi.get(exportUrl, {
      baseURL: '', // unset baseUrl since we're not hitting the normal sheets API
      params: {
        id: this.spreadsheetId,
        format: fileType,
        ...worksheetId && { gid: worksheetId },
      },
      responseType: returnStreamInsteadOfBuffer ? 'stream' : 'arraybuffer',
    });
    return response.data;
  }

  /** exports entire document as html file (zipped) */
  async downloadAsHTML(returnStreamInsteadOfBuffer = false) {
    return this._downloadAs('html', undefined, returnStreamInsteadOfBuffer);
  }
  /** exports entire document as xlsx spreadsheet (Microsoft Office Excel) */
  async downloadAsXLSX(returnStreamInsteadOfBuffer = false) {
    return this._downloadAs('xlsx', undefined, returnStreamInsteadOfBuffer);
  }
  /** exports entire document as ods spreadsheet (Open Office) */
  async downloadAsODS(returnStreamInsteadOfBuffer = false) {
    return this._downloadAs('ods', undefined, returnStreamInsteadOfBuffer);
  }


  async delete() {
    const response = await this.driveApi.delete('');
    this._deleted = true;
    return response.data;
  }

  // PERMISSIONS ///////////////////////////////////////////////////////////////////////////////////
  async listPermissions() {
    const listReq = await this.driveApi.request({
      method: 'GET',
      url: '/permissions',
      params: {
        fields: 'permissions(id,type,emailAddress,domain,role,displayName,photoLink,deleted)',
      }
    });
    return listReq.data.permissions as PermissionsList;
  }

  async setPublicAccessLevel(role: PublicPermissionRoles | false) {
    const permissions = await this.listPermissions();
    const existingPublicPermission = _.find(permissions, (p) => p.type === 'anyone');
    
    if (role === false) {
      if (!existingPublicPermission) {
        // doc is already not public... could throw an error or just do nothing
        return;
      }
      await this.driveApi.request({
        method: 'DELETE',
        url: `/permissions/${existingPublicPermission.id}`,
      });
    } else {
      const shareReq = await this.driveApi.request({
        method: 'POST',
        url: '/permissions',
        params: {
        },
        data: {
          role: role || 'viewer',
          type: 'anyone',
        }
      });
    }
  }
  
  async share(emailAddressOrDomain: string, opts?: {
    /** set role level, defaults to owner */
    role?: PermissionRoles,

    /** set to true if email is for a group */
    isGroup?: boolean,
    
    /** set to string to include a custom message, set to false to skip sending a notification altogether */
    emailMessage?: string | false,

    // moveToNewOwnersRoot?: string,
    // /** send a notification email (default = true) */
    // sendNotificationEmail?: boolean,
    // /** support My Drives and shared drives (default = false) */
    // supportsAllDrives?: boolean,
    
    // /** Issue the request as a domain administrator */
    // useDomainAdminAccess?: boolean,
  }) {
    let emailAddress: string | undefined;
    let domain: string | undefined;
    if (emailAddressOrDomain.includes('@')) {
      emailAddress = emailAddressOrDomain;
    } else {
      domain = emailAddressOrDomain;
    }


    const shareReq = await this.driveApi.request({
      method: 'POST',
      url: '/permissions',
      params: {
        ...opts?.emailMessage === false && { sendNotificationEmail: false },
        ..._.isString(opts?.emailMessage) && { emailMessage: opts?.emailMessage },
        ...opts?.role === 'owner' && { transferOwnership: true },
      },
      data: {
        role: opts?.role || 'writer',
        ...emailAddress && {
          type: opts?.isGroup ? 'group' : 'user',
          emailAddress,
        },
        ...domain && {
          type: 'domain',
          domain,
        }
      }
    });

    return shareReq.data;
  }

  //
  // CREATE NEW DOC ////////////////////////////////////////////////////////////////////////////////
  static async createNewSpreadsheetDocument(auth: GoogleSpreadsheetAuth, properties?: Partial<SpreadsheetProperties>) {
    // see updateProperties for more info about available properties
    
    if ('apiKey' in auth) {
      throw new Error('Cannot use api key only to create a new spreadsheet - it is only usable for read-only access of public docs')
    }

    // TODO: handle injecting default credentials if running on google infra 

    const authConfig = await getRequestAuthConfig(auth);

    const response = await Axios.request({
      method: 'POST',
      url: SHEETS_API_BASE_URL,
      paramsSerializer: axiosParamsSerializer,
      ...authConfig, // has the auth header
      data: {
        properties
      }
    });

    const newSpreadsheet = new GoogleSpreadsheet(response.data.spreadsheetId, auth);

    // TODO ideally these things aren't public, might want to refactor anyway
    newSpreadsheet._spreadsheetUrl = response.data.spreadsheetUrl;
    newSpreadsheet._rawProperties = response.data.properties;
    _.each(response.data.sheets, (s) => newSpreadsheet._updateOrCreateSheet(s));

    return newSpreadsheet;  
  }
}
