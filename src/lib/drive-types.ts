// export type PermissionRoles = 'owner' | 'organizer' | 'fileOrganizer' | 'writer' | 'commenter' | 'reader';
export type PermissionRoles = 'owner' | 'writer' | 'commenter' | 'reader';
export type PublicPermissionRoles = Exclude<PermissionRoles, 'owner'>;

// this shape is set by what we request...
type PublicPermissionListEntry = {
  id: 'anyoneWithLink'
  type: 'anyone',
  role: PublicPermissionRoles,
}
type UserOrGroupPermissionListEntry = {
  id: string;
  displayName: string;
  type: 'user' | 'group';
  photoLink?: string;
  emailAddress: string;
  role: PermissionRoles;
  deleted: boolean;
}
type DomainPermissionListEntry = {
  id: string;
  displayName: string;
  type: 'domain';
  domain: string;
  role: PublicPermissionRoles;
  photoLink?: string;
}

export type PermissionsList = (PublicPermissionListEntry | UserOrGroupPermissionListEntry | DomainPermissionListEntry)[];