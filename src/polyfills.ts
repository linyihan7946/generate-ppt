if (!(Object as any).hasOwn) {
    (Object as any).hasOwn = (target: object, key: PropertyKey): boolean =>
        Object.prototype.hasOwnProperty.call(target, key);
}
