# Todo

## Phase 1 - Core API
- [x] Project setup with uv and openpyxl
- [x] Implement `set_border()` for range-based border application
- [x] Unit tests for single cell and range borders
- [ ] Support `xlEdgeLeft`, `xlEdgeRight` etc. VBA-style border index constants
- [ ] `set_outline()` convenience function (outer borders only)

## Phase 2 - Advanced Features
- [ ] Merge-cell aware border application
- [ ] Named style presets (e.g. "box", "grid", "outline")
- [ ] Color support for border sides
- [ ] Diagonal border support

## Phase 3 - Polish
- [ ] README with usage examples
- [ ] Type stub / py.typed marker
- [ ] Publish to PyPI
