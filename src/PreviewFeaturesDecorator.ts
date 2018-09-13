/**
 * Decorator marking that the class (command) is using preview or beta functionality
 * @param target the class
 */
export const Preview: ClassDecorator = (target) => {
  target.prototype.previewFeaturesUsed = true;
  return target;
}
