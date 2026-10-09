# Real-world test corpus — sources

**Test use only.** These files exercise pypptx against decks saved by PowerPoint. They are
not part of the library: packages are found under `src/` only, so nothing here reaches the
wheel. Do not use them in `examples/` or the docs.

All files come unmodified from Apache POI's test data, at SVN revision 1938723 of
`https://svn.apache.org/repos/asf/poi/trunk/test-data/slideshow/`, under the Apache License,
Version 2.0 (`LICENSE` here; attribution in `NOTICE`). Selection rule: POI's own
bug-reproduction and sample files only; decks POI collected from public websites are
excluded. Fetch one again with:

```
curl -O 'https://svn.apache.org/repos/asf/!svn/bc/1938723/poi/trunk/test-data/slideshow/<file>'
```

Authoring application is read from each file's `docProps/app.xml`. `2411-Performance_Up`, `bug68703`,
`customGeo` and `EmbeddedVideo` were added in v0.5.0 (same revision) for transitions and animations.

| File | KB | Saved by | Exercises | SHA-256 |
|---|---|---|---|---|
| `2411-Performance_Up.pptx` | 633 | Microsoft PowerPoint 12.0001 | transition, animation (after-previous, by paragraph), 48 slides | `a387861a0066ad11ba85aba9572830c1ec1ab9a24380728ead34d2b5b26662c0` |
| `45545_Comment.pptx` | 301 | Microsoft PowerPoint 7.0 12.0000 | comments, notes, transition, animation, 11 slides | `0295a51f63150e3ee680154d07dab88062ba2306030b790561f85436f3ded7c8` |
| `54542_cropped_bitmap.pptx` | 97 | Microsoft Office PowerPoint 14.0000 | picture, group, 1 slide | `fa219615dae3cd01f62e8449f7b477bebadd53d3cfba3348ada6377098d8c65e` |
| `backgrounds.pptx` | 61 | Microsoft Office PowerPoint 14.0000 | background, 4 slides | `6d1e661c87b072a701e02e28f9d2c9bb43c6bfe0ad5acd02811a0382bab769b7` |
| `bar-chart.pptx` | 43 | Microsoft Macintosh PowerPoint 15.0033 | chart, 1 slide | `79e1d218bfb2903e8dc8425a6b1997d9c1976f5a5f025bada85b0c47b5777969` |
| `bug58144-headers-footers-2007.pptx` | 41 | Microsoft Office PowerPoint 16.0000 | notes, 1 slide | `63e0db239ef3083b8537705941247af195b200e83535f8d92c246e774f771289` |
| `bug60715.pptx` | 72 | Microsoft Office PowerPoint 16.0000 | transition, 1 slide | `d560102b417036323dd21781b13e86b326e3d4d13ece3419aaf7c55a9c22b8d4` |
| `bug60993.pptx` | 21 | LibreOffice/5.3.2.2$Linux_X86_64 LibreOffice_project/30m0$Build-2 | table, animation, 1 slide | `ad54b71b1deeec29d9e5e6d8d81ae4590b5ffafb906ea738aebe347336189364` |
| `bug65523.pptx` | 49 | Microsoft Office PowerPoint 14.0000 | media, picture, 1 slide | `a37c15125aa21ac5226bac07f4459a200cf8070040c06093c1dc1db19e28c12c` |
| `bug68703.pptx` | 59 | Microsoft Macintosh PowerPoint 16.0000 | transition in `mc:AlternateContent` (p14), animation, 1 slide | `f232d0dc390c741855b4395f74a11d2610243fffd8699f72dd3185ad5f930852` |
| `chart-slide-bg.pptx` | 55 | Microsoft Office PowerPoint 16.0000 | chart, background, 1 slide | `1e57296fdac894c239ed47e4fc9521da59ab26c34f4d28ec2036f0befbc5fd74` |
| `customGeo.pptx` | 1022 | Microsoft Office PowerPoint 14.0000 | transition, animation (click, with-previous, after-previous, by paragraph), 48 slides | `9baa7f4554cbaddfcebcc431432268ea34f20c02787172bb0a7d12d640e7e84f` |
| `EmbeddedAudio.pptx` | 87 | Microsoft Macintosh PowerPoint 16.0000 | media, picture, animation, 1 slide | `c5ae4274e2bf5504a56aef9c8d7c5d2381ece69c9bf68f7749ad5eae3e675edb` |
| `EmbeddedVideo.pptx` | 197 | Microsoft Office PowerPoint 16.0000 | media, animation (triggered sequence), 1 slide | `7940e3b1a339db11f00b65399a2fe77e0e85a5da3a30ac8d6c8a0a77527b2ab2` |
| `layouts.pptx` | 61 | Microsoft Office PowerPoint 14.0000 | picture, notes, 10 slides | `9c3d53afa3115de2ba72ec77637a616aeed505a6ff3b30a46db871c4edd32be5` |
| `line-chart.pptx` | 43 | Microsoft Macintosh PowerPoint 15.0033 | chart, 1 slide | `a2319540fb096629874e8c2baf91b9f8afd1386bfba411efff503b96dce9e9a1` |
| `missing-blip-fill.pptx` | 39 | Microsoft Macintosh PowerPoint 16.0000 | picture, notes, transition, background, 1 slide | `00b19b1ba0f3b399dd3c3c0de906b8cfeca5cda9425d481c2d75f5d74790c004` |
| `pie-chart.pptx` | 54 | Microsoft Office PowerPoint 14.0000 | chart, 1 slide | `3b6404b59b24cb79fbb91fc2e92bd8b80cdc340aeed71a1ec1e267db0d8ad444` |
| `prProps.pptx` | 30 | Microsoft Office PowerPoint 14.0000 | notes, animation, 1 slide | `b6d031d7cd670a007489465c57a0f233c497e89aed62a0a9f7d916cf173c86ad` |
| `radar-chart.pptx` | 42 | Microsoft Macintosh PowerPoint 15.0033 | chart, 1 slide | `faf631199e44a1eb5dae48527d233b3231e67ea0345e93fd0c9306ce5261b570` |
| `sample_pptx_grouping_issues.pptx` | 38 | Microsoft Office PowerPoint 12.0000 | group, notes, animation, 1 slide | `a8639e12e763a01d889a6178e154c21db2dd4275f8f35ac81a458f81e26b2c18` |
| `scatter-chart.pptx` | 43 | Microsoft Macintosh PowerPoint 15.0033 | chart, 1 slide | `946d17760f23822950d97a02aeeedbac773fe8ba5afd61bb26a9ee5fd8db53be` |
| `smartart-simple.pptx` | 69 | Microsoft Macintosh PowerPoint 16.0000 | SmartArt, 1 slide | `238ff6433dac00b2722edb9a29d97f56ab4f3c9c37f33a73d6f53baf494916aa` |
| `SmartArt.pptx` | 40 | Microsoft Office PowerPoint 16.0000 | SmartArt, 1 slide | `b97e4c6d2ee1dd4094f50f9043820610268dd452d7717c537f5456636adcc353` |
| `table-with-theme.pptx` | 34 | Microsoft Macintosh PowerPoint 16.0000 | table, 1 slide | `571d0cd2e1e9767249b3a02544c7771c883c48bb191264409729c7cec574ab48` |
| `table_test2.pptx` | 32 | Microsoft Office PowerPoint 16.0000 | table, 1 slide | `e80d09a5fa767f6ee2d024f3a4198f1e8b3241cc638183e5ebafdbacf1fb2838` |
| `with_japanese.pptx` | 46 | Microsoft Office PowerPoint 12.0000 | table, animation, 1 slide | `f0b1efbfbd70f7ee2395dc2505639f99449971ba404a6ec8cf1b97d82cf45c10` |
| `WithMaster.pptx` | 46 | Microsoft Office PowerPoint 14.0000 | 2 slides | `54ae7677fd2813a815160cc72bd8960a4e482690dfc369b392868521c9b61549` |

## `malformed/`

Fuzzer-minimized files from OSS-Fuzz (ClusterFuzz), from the same POI directory and revision.
None is a valid zip package; each must fail with a clear pypptx exception.

| File | KB | SHA-256 |
|---|---|---|
| `clusterfuzz-testcase-minimized-POIFuzzer-5205835528404992.pptx` | 35 | `f568b01439e5012760839f4d0d197dba99c57ca72a66b277515428fb07e0e136` |
| `clusterfuzz-testcase-minimized-POIXSLFFuzzer-4838644450394112.pptx` | 22 | `b09ec6479441ec761a3736c09453a6c6cc8ad413c571e738d3810b8338893270` |
| `clusterfuzz-testcase-minimized-POIXSLFFuzzer-5611274456596480.pptx` | 26 | `e098775305cc81413b8a53a3c28df3862cd73ffc4eacf2395649bef0afc0d1ec` |
| `clusterfuzz-testcase-minimized-POIXSLFFuzzer-6254434927378432.pptx` | 7 | `27fb0d0641003899ab36e96e1aa1811eac43c5e1ac496a148b9d1f217db31ffc` |
