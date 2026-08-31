# Releasing

`python-pptx-op` is published to PyPI by `.github/workflows/release.yml`, which runs on a
`v*` tag. Releases use **trusted publishing** (OpenID Connect), so no PyPI API token is
stored in the repository.

## One-time setup

Both steps are required before the first release, and only need doing once.

### 1. Register the trusted publisher on PyPI

The project does not exist on PyPI yet, so register it as a **pending publisher** at
<https://pypi.org/manage/account/publishing/>:

| Field | Value |
| --- | --- |
| PyPI project name | `python-pptx-op` |
| Owner | `vitordj` |
| Repository name | `python-pptx-op` |
| Workflow name | `release.yml` |
| Environment name | `pypi` |

A pending publisher becomes a normal one automatically when the first release is
published under it.

### 2. Create the `pypi` GitHub environment

In the repository settings, under *Environments*, create one named `pypi`. The workflow
references it, and it is the natural place to add required reviewers if you want a manual
approval gate before anything reaches PyPI.

## Cutting a release

1. Update `HISTORY.rst`: change the top heading from `X.Y.Z (unreleased)` to
   `X.Y.Z (YYYY-MM-DD)` and confirm the entries match what actually changed.
2. Bump `__version__` in `src/pptx/__init__.py`. This is the single source of truth —
   `pyproject.toml` reads it via `[tool.setuptools.dynamic]`, and the release workflow
   fails if the tag does not match it.
3. Verify locally (see below).
4. Commit, then tag and push:

   ```bash
   git commit -am "release: prepare vX.Y.Z"
   git tag vX.Y.Z
   git push origin master
   git push origin vX.Y.Z
   ```

5. Watch the `release` workflow. It builds the sdist and wheel, runs
   `twine check --strict`, verifies the tag against `__version__`, installs the built
   wheel into a clean virtualenv and smoke-tests it, then publishes.

Running the workflow manually (*Run workflow* / `workflow_dispatch`) performs every build
and check step but **does not publish** — useful for validating a release before tagging.

## Verifying locally

What the workflow does, reproduced on your machine:

```bash
rm -rf dist build *.egg-info
python -m build                      # sdist + wheel
python -m twine check --strict dist/*

python -m venv /tmp/relcheck
/tmp/relcheck/bin/pip install dist/*.whl
/tmp/relcheck/bin/python -c "import pptx; print(pptx.__version__)"
```

Run the test suite and build the docs too:

```bash
pytest                               # 2795 tests
make -C docs html                    # must be warning-free
```

## Publishing a test release

To rehearse against TestPyPI, register a second pending publisher there with the same
values, then add a TestPyPI step to the workflow or upload manually:

```bash
python -m twine upload --repository testpypi dist/*
```

Note that a version number can never be reused on PyPI, even after deleting a release, so
rehearse with a `.devN` or `.rcN` suffix rather than the real version.

## Version numbering

The fork keeps upstream's version line as its base: it forked at python-pptx 1.0.2, and
its own releases start at 1.1.0. Additions bump the minor version; fixes bump the patch
version.
