# Use Biome and architecture tests for code style

The project will use Biome for formatting and baseline linting, with Vitest architecture tests enforcing conventions that Biome does not express well: App Router colocation under underscore directories and the "no new application classes" rule. This keeps the fast default style tool the team chose while still making Next.js App Router placement and function-style conventions fail automatically when AI or human changes drift.
