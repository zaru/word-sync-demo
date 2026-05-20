# Copilot coding instructions

Follow these project conventions for every code change.

## Next.js App Router and React 19

- Prefer Server Components. Add `use client` only to the smallest component that needs hooks, browser APIs, or event handlers.
- Keep data loading, cookie access, auth boundaries, and persistence on the server side unless browser execution is required.
- Use Route Handlers for public HTTP boundaries, browser polling/autosave endpoints, and external callbacks.
- Prefer Server Actions in `_actions` for UI-originated server mutations such as form submissions.

## App Router file placement

- In `src/app`, route segment roots should contain only Next.js special files such as `page.tsx`, `layout.tsx`, `route.ts`, `loading.tsx`, `error.tsx`, `not-found.tsx`, metadata files, route tests, and the root `globals.css`.
- Put route-local UI in `_components`.
- Put route-local Server Actions in `_actions`.
- Put route-local helpers, clients, and non-UI utilities in `_lib`.
- Keep tests beside the file they test, including `route.test.ts` beside `route.ts`.
- Keep shared domain logic and shared server boundaries outside `src/app`, under `src/domain` and `src/server`.

## Programming style

- Prefer function style: functions, factory functions, plain objects, and closures.
- Do not add new application `class` declarations unless an external library API or a clearly documented interoperability boundary requires one.
- Existing store classes are legacy exceptions. When touching them for substantial changes, prefer moving them toward factory functions rather than adding more classes.
- Preserve strict TypeScript types. Avoid `any` and broad type assertions; use narrow guards and explicit errors.

## Checks

- Use pnpm.
- Run `pnpm lint` for Biome format/lint checks.
- Run `pnpm test` for behavior and architecture tests.
