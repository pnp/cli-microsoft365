/**
 * Specifies a set of built-in permissions.
 * See: https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.basepermissions.aspx
 */
export class BasePermissions {
  private _high: number = 0;
  private _low: number = 0;

  public get high(): number {
    return this._high;
  }

  public get low(): number {
    return this._low;
  }

  public set high(value:number) {
    this._high = value;
  }

  public set low(value:number) {
    this._low = value;
  }

  public has(perm: PermissionKind) : boolean {
    var hasPermission = false;

    if (perm === PermissionKind.FullMask)
    {
      hasPermission = (this.high & 32767) === 32767 && this.low === 65535;
    }

    var a = perm;
    a = a - 1;
    var b = 1;
    if (a >= 0 && a < 32) {
      b = b << a;
      hasPermission = 0 !== (this.low & b);
    } else if (a >= 32 && a < 64) {
      b = b << a - 32;
      hasPermission = (0 !== (this.high & b));
    }
    
    return hasPermission;
  }

  public set(perm: PermissionKind): void {
    if (perm == PermissionKind.FullMask) {
      this._low = 65535;
      this._high = 32767;
    }
    else if (perm == PermissionKind.EmptyMask) {
      this._low = 0;
      this._high = 0;
    }
    else {
      let num1: number = (perm - 1);
      let num2: number = 1;
      if (num1 >= 0 && num1 < 32) {
        this._low = this._low | num2 << num1;
      }
      else {
        if (num1 < 32 || num1 >= 64) {
          return;
        }

        this._high = this._high | num2 << num1 - 32;
      }
    }
  }
}

/**
 * Specifies permissions that are used to define user roles.
 * See: https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.permissionkind.aspx
 */
export enum PermissionKind {
  EmptyMask = 0,
  ViewListItems = 1,
  AddListItems = 2,
  EditListItems = 3,
  DeleteListItems = 4,
  ApproveItems = 5,
  OpenItems = 6,
  ViewVersions = 7,
  DeleteVersions = 8,
  CancelCheckout = 9,
  ManagePersonalViews = 10,
  ManageLists = 12,
  ViewFormPages = 13,
  AnonymousSearchAccessList = 14,
  Open = 17,
  ViewPages = 18,
  AddAndCustomizePages = 19,
  ApplyThemeAndBorder = 20,
  ApplyStyleSheets = 21,
  ViewUsageData = 22,
  CreateSSCSite = 23,
  ManageSubwebs = 24,
  CreateGroups = 25,
  ManagePermissions = 26,
  BrowseDirectories = 27,
  BrowseUserInfo = 28,
  AddDelPrivateWebParts = 29,
  UpdatePersonalWebParts = 30,
  ManageWeb = 31,
  AnonymousSearchAccessWebLists = 32,
  UseClientIntegration = 37,
  UseRemoteAPIs = 38,
  ManageAlerts = 39,
  CreateAlerts = 40,
  EditMyUserInfo = 41,
  EnumeratePermissions = 63,
  FullMask = 65
}