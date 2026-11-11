# Releasing to PyPI

The repository is prepared for `docx-typst` version `0.1.0`. Preparing artifacts
does not publish them. Perform release validation from a clean checkout.

## Validate the candidate

```console
python -m pip install -e ".[test,release]"
python -m pytest -q
python -m build
python -m twine check --strict dist/*
python tools/check_release.py dist
```

Use a clean `dist` directory containing only the candidate's wheel and source
archive. `python -m build` builds the wheel from the source archive, checking that
the source distribution is sufficient to build the package. The release checker
checks version consistency, packaged resources, and source documentation/skills.

Install the built wheel into a fresh environment, with its dependencies, and run
`tools/smoke_release.py` using that environment's Python. Do not use an editable
install or expose the checkout through `PYTHONPATH`. The smoke check exercises
installed CLI entry points, packaged starter resources, calculated DOCX export,
and exact source roundtrip. CI performs this check before publishing.

The local test result does not establish support on every platform. Review the
GitHub Actions Python 3.10–3.13 matrix on Linux, Windows, and macOS before tagging.
Word visual checks are separate from headless conversion tests.

## Configure trusted publishing

In PyPI, configure the project's GitHub Trusted Publisher with these values:

| Setting | Value |
|---|---|
| Project | `docx-typst` |
| Owner | `Sovichea` |
| Repository | `docx-typst` |
| Workflow filename | `publish-pypi.yml` |
| Environment | `pypi` |

For a first release, use a pending publisher. Create the matching `pypi` GitHub
environment and configure any desired release reviewers. Repository configuration
alone does not create or verify the PyPI publisher or GitHub environment.

The workflow uses job-scoped `id-token: write` and PyPA's publishing action;
no PyPI API token is needed. See the official
[Trusted Publishing instructions](https://docs.pypi.org/trusted-publishers/using-a-publisher/).

## Publish an approved release

Update both `pyproject.toml` and `src/docx_typst/__init__.py` when changing the
version, update the changelog, and repeat validation. PyPI versions cannot be
replaced by re-uploading different files under the same version.

After release approval and successful CI, create and push the matching tag:

```console
git tag -a v0.1.0 -m "Release 0.1.0"
git push origin v0.1.0
```

Pushing a `v*` tag triggers publication. The workflow rejects version mismatches,
tests the candidate, builds and checks both distributions, smoke-tests the wheel
in a clean environment, and publishes the same uploaded artifacts. Check the
workflow and PyPI result, then verify installation from PyPI in a fresh environment.
