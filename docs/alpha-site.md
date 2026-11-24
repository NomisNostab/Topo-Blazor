# Alpha site

Every push to `develop` publishes an **alpha** copy of Topo, and (once AWS is set up) an alpha copy of the report function, so changes can be tried live before a release is cut from `main`.

| | Live | Alpha |
|---|---|---|
| Branch | `main` | `develop` |
| Site | `https://nomisnostab.github.io/Topo-Blazor/` | `https://nomisnostab.github.io/Topo-Blazor/alpha/` |
| Report function | the live Lambda | an alpha Lambda (or the live one until it exists) |
| Workflow | `.github/workflows/main.yml` | `.github/workflows/alpha.yml` |

The alpha site is a folder (`alpha/`) of the same `gh-pages` branch the live site is served from, because a repository has one GitHub Pages site. It has a red **ALPHA - develop abc1234 - not released** tag in the corner, so nobody mistakes it for the live site, and it says which commit it is.

**The alpha site signs in to the real Terrain.** Anything on `develop` that writes to Terrain writes real records when used on the alpha site. Reports made there carry real members' names through the alpha Lambda.

## What the workflow does

1. Runs the unit tests. A failing test stops the deploy.
2. If `ALPHA_LAMBDA_FUNCTION` is set: publishes `TopoReportFunction` (linux-x64, about 28 MB zipped) and updates that function with `aws lambda update-function-code`. This comes before the site so the site never talks to a function that does not know its report types.
3. Replaces the Syncfusion key (`SYNCFUSION_KEY`, as `main.yml` does) and, if `ALPHA_REPORT_URL` is set, the report function address in `ReportService.cs`.
4. Publishes the site, sets its base path to `/Topo-Blazor/alpha/`, and adds the banner.
5. Publishes it to the `alpha` folder of `gh-pages`. Only that folder is replaced.

It runs on pushes to `develop` and by hand (Actions > Deploy alpha > Run workflow). It does not run for pull requests, so code from a fork never runs with the secrets.

Until the AWS settings below exist, steps 2 and the address swap are skipped and the alpha site makes its reports with the live Lambda. That works, but a report type that only exists on `develop` will fail until the alpha Lambda is set up.

## One-time setup

Repository settings (Settings > Secrets and variables > Actions):

| Name | Kind | Value |
|---|---|---|
| `SYNCFUSION_KEY` | secret | already set for `main.yml` |
| `ALPHA_AWS_ROLE_ARN` | secret | ARN of the role GitHub assumes (below) |
| `ALPHA_LAMBDA_FUNCTION` | variable | the alpha function's name, e.g. `TopoReportGeneratorAlpha` |
| `ALPHA_REPORT_URL` | variable | the alpha function's Function URL, with the trailing `/` |

In AWS (region `ap-southeast-2`):

1. **The alpha function.** Same settings as the live one: runtime .NET 10, x86_64, 1024 MB, 300 s timeout, handler `TopoReportFunction::TopoReportFunction.Function::FunctionHandler`, same execution role. Add the environment variable `SyncfusionLicense` (the key, so PDFs have no licence watermark). The first code can be anything; the workflow replaces it.
2. **A Function URL** for it with auth type `NONE` and CORS allowing the origin `https://nomisnostab.github.io`, methods `PUT` and `POST`, header `content-type`. Put that URL in `ALPHA_REPORT_URL`.
3. **Limits.** A Function URL with no auth can be called by anyone who has the address. Set reserved concurrency on the function (for example 5) and a billing alert.
4. **GitHub sign-in without stored keys.** Add the identity provider `token.actions.githubusercontent.com` (audience `sts.amazonaws.com`) if the account does not have it, then a role GitHub can assume, with a trust condition limited to this repository's `develop` branch:

   ```json
   {
     "Effect": "Allow",
     "Principal": { "Federated": "arn:aws:iam::<account>:oidc-provider/token.actions.githubusercontent.com" },
     "Action": "sts:AssumeRoleWithWebIdentity",
     "Condition": {
       "StringEquals": {
         "token.actions.githubusercontent.com:aud": "sts.amazonaws.com",
         "token.actions.githubusercontent.com:sub": "repo:NomisNostab/Topo-Blazor:ref:refs/heads/develop"
       }
     }
   }
   ```

   Its permissions need only `lambda:UpdateFunctionCode`, `lambda:GetFunction` and `lambda:GetFunctionConfiguration` on the alpha function.

## Things to know

- **Open the alpha site at its folder address.** GitHub Pages serves one `404.html`, the one at the site root, which belongs to the live site. A link straight to a page inside `alpha/` (for example `/alpha/members`) will not load the alpha app; start from `/alpha/`.
- **`main.yml` must not wipe the alpha folder.** It uses version 3.7.1 of the deploy action, which does not delete files it did not write. After the first release following this change, check that `/alpha/` still loads. If it does not, exclude the `alpha` folder from cleaning in `main.yml` (the action's `CLEAN_EXCLUDE` option) or update that action. This is from the action's documented defaults and has not been tried on this repository.
- **Both workflows push to `gh-pages`.** `alpha.yml` queues behind itself with a `concurrency` group; giving `main.yml` the same group (`gh-pages-deploy`) would stop the two ever pushing at once.
- The `.github` change reaches `main` with the next release; until then only `develop` has it, and a workflow only runs for the branch it is on, which is what is wanted.
