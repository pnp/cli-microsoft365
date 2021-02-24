export function getControlTypeDisplayName(controlType: number): string {
  switch (controlType) {
    case 0:
      return 'Empty column';
    case 3:
      return 'Client-side web part';
    case 4:
      return 'Client-side text';
    default:
      return '' + controlType;
  }
}