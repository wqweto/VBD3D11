## DirectX 11 for VB6 1.0 Type Library

A fairly complete VB6-compatible DirectX 11 type library

### Description

This project is a work-in-progress on bringing D3D, D3D11 and DXGI APIs incl. enums, structs, interfaces and functions to VB6.

Some of the declarations are left as stubs yet, with some interfaces aliased to `IUnknown` and structs declared as `void`.

### Usage

In VB6 IDE just add reference to `VBD3D11.tlb` in the `typelib` directory.

The `VBD3D11.tlb` file is needed only in VB6 IDE and there is no need to ship it with final executables (please don't).

### Samples

Check out the `tutorials` directory for code samples.
