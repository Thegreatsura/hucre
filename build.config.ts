import { defineBuildConfig } from "obuild/config";

export default defineBuildConfig({
  entries: [
    {
      type: "bundle",
      input: ["./src/index.ts", "./src/xlsx.ts", "./src/csv.ts", "./src/ods.ts", "./src/cli.ts"],
      minify: true,
    },
  ],
});
