import * as assert from 'assert';
import { BasePermissions, PermissionKind } from './base-permissions'

describe('BasePermissions', () => {

  let basePermissions: BasePermissions;
  const getPermissions = (rights: PermissionKind[]): BasePermissions => {
    for (let kind of rights) {
        basePermissions.set(kind);
    }
    return basePermissions;
  }

  beforeEach(() => {
    basePermissions = new BasePermissions();
  });

  it('has correct permissions set with AddListItems and DeleteListItems', () => {
    const delegatedPermissions: PermissionKind[] = [
      PermissionKind.AddListItems, PermissionKind.DeleteListItems
    ];
    const result: BasePermissions = getPermissions(delegatedPermissions);

    assert.strictEqual(result.low, 10);
    assert.strictEqual(result.high, 0);
  });

  it('has correct permissions set with ManageLists, AddListItems and DeleteListItems', () => {
    const delegatedPermissions: PermissionKind[] = [
      PermissionKind.ManageLists, PermissionKind.AddListItems, 
      PermissionKind.DeleteListItems
    ];
    const result: BasePermissions = getPermissions(delegatedPermissions);

    assert.strictEqual(result.low, 2058);
    assert.strictEqual(result.high, 0);
  });

  it('has correct permissions set with ManageLists', () => {
    const delegatedPermissions: PermissionKind[] = [
      PermissionKind.ManageLists
    ];
    const result: BasePermissions = getPermissions(delegatedPermissions);

    assert.strictEqual(result.low, 2048);
    assert.strictEqual(result.high, 0);
  });

  it('has correct permissions set with FullMask', () => {
    const delegatedPermissions: PermissionKind[] = [
      PermissionKind.FullMask
    ];
    const result: BasePermissions = getPermissions(delegatedPermissions);

    assert.strictEqual(result.low, 65535);
    assert.strictEqual(result.high, 32767);
  });

  it('has correct permissions set with EmptyMask', () => {
    const delegatedPermissions: PermissionKind[] = [
      PermissionKind.EmptyMask
    ];
    const result: BasePermissions = getPermissions(delegatedPermissions);

    assert.strictEqual(result.low, 0);
    assert.strictEqual(result.high, 0);
  });

  it('has correct permissions set with EmptyMask and AddListItems', () => {
    const delegatedPermissions: PermissionKind[] = [
      PermissionKind.EmptyMask, PermissionKind.AddListItems
    ];
    const result: BasePermissions = getPermissions(delegatedPermissions);

    assert.strictEqual(result.low, 2);
    assert.strictEqual(result.high, 0);
  });

  it('has correct permissions set with AddListItems, DeleteListItems and FullMask', () => {
    const delegatedPermissions: PermissionKind[] = [
      PermissionKind.AddListItems, PermissionKind.DeleteListItems,
      PermissionKind.FullMask
    ];
    const result: BasePermissions = getPermissions(delegatedPermissions);

    assert.strictEqual(result.low, 65535);
    assert.strictEqual(result.high, 32767);
  });

  it('has correct permissions set with ManagePermissions', () => {
    const delegatedPermissions: PermissionKind[] = [
      PermissionKind.ManagePermissions
    ];
    const result: BasePermissions = getPermissions(delegatedPermissions);

    assert.strictEqual(result.low, 33554432);
    assert.strictEqual(result.high, 0);
  });

  it('has correct permissions set with ManageWeb', () => {
    const delegatedPermissions: PermissionKind[] = [
      PermissionKind.ManageWeb
    ];
    const result: BasePermissions = getPermissions(delegatedPermissions);

    assert.strictEqual(result.low, 1073741824);
    assert.strictEqual(result.high, 0);
  });

  it('has correct permissions set with ManageWeb and FullMask', () => {
    const delegatedPermissions: PermissionKind[] = [
      PermissionKind.ManageWeb, PermissionKind.FullMask
    ];
    const result: BasePermissions = getPermissions(delegatedPermissions);

    assert.strictEqual(result.low, 65535);
    assert.strictEqual(result.high, 32767);
  });

  it('has correct permissions set with EnumeratePermissions', () => {
    const delegatedPermissions: PermissionKind[] = [
      PermissionKind.EnumeratePermissions
    ];
    const result: BasePermissions = getPermissions(delegatedPermissions);

    assert.strictEqual(result.low, 0);
    assert.strictEqual(result.high, 1073741824);
  });

  it('exits correctly on incorrect permissions set', () => {
    basePermissions.set((-1 as PermissionKind));

    assert.strictEqual(basePermissions.low, 0);
    assert.strictEqual(basePermissions.high, 0);
  });

  it('has correct high and low value set', () => {
    basePermissions.high = 32767;
    basePermissions.low = 65535;

    assert.strictEqual(basePermissions.high, 32767);
    assert.strictEqual(basePermissions.low, 65535);
  });

  it('checks the permission correctly for the actual FullMask', () => {

    //http://aaclage.blogspot.co.uk/2014/09/it-is-sharepoint-permission-call.html    
    // Full permission.
    basePermissions.high = 2147483647;
    basePermissions.low = 4294967295;

    assert.strictEqual(basePermissions.has(PermissionKind.AddAndCustomizePages), true);
    assert.strictEqual(basePermissions.has(PermissionKind.ManageWeb), true);
    assert.strictEqual(basePermissions.has(PermissionKind.EnumeratePermissions), true);
  });

  it('checks the permission correctly for the online FullMask', () => {
    
    // Full permission.
    basePermissions.high = 32767;
    basePermissions.low = 65535;

    assert.strictEqual(basePermissions.has(PermissionKind.FullMask), true);
  });
});