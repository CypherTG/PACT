// Ambient type declarations for CSS Modules used in this SPFx project.
// SPFx webpack handles the actual CSS Module class-name transformation at
// build time. This file exists solely to satisfy the TypeScript compiler
// when running tsc --noEmit or IDE type-checking.
//
// NOTE: This file must NOT contain any top-level import/export statements
// (it must remain a "script" file, not a "module") for the wildcard declare
// module to work as a global ambient declaration.

declare module '*.module.css' {
  const styles: { readonly [className: string]: string };
  export default styles;
}
