import js from '@eslint/js';
import tseslint from 'typescript-eslint';
import prettierConfig from 'eslint-config-prettier';
import globals from 'globals';

export default tseslint.config(
  {
    ignores: [
      '**/dist/**',
      '**/node_modules/**',
      '**/.turbo/**',
      '**/coverage/**',
      '.claude/**',
      // 验证/方案验证脚本 + 已废弃归档：不参与功能行为，不纳入 lint 门禁
      // （与 vitest.config.ts 对 dot-extractor 的排除一致；决议见 CLAUDE.md「废弃模块」）
      'examples/skeleton-cli/**',
      'packages/dot-extractor/**',
      '_deprecated-schemes/**',
      '.spike/**', // 方案验证 spike（已 gitignore，不入库）
      'coworkcoco-officeexamplesskeleton-clim5_e2e_test.mts', // 路径拼接事故畸形文件（已 gitignore）
    ],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
      globals: {
        ...globals.node,
        ...globals.browser,
        ...globals.worker,
      },
    },
    rules: {
      '@typescript-eslint/no-unused-vars': [
        'error',
        { argsIgnorePattern: '^_', varsIgnorePattern: '^_' },
      ],
      '@typescript-eslint/consistent-type-imports': 'error',
      'no-console': ['warn', { allow: ['warn', 'error'] }],
      // 全角空格 `　`(U+3000)在本库是**合法数据/文档**：公文模板/印发行(`{印发机关}　{印发日期}`)、
      //   喂 LLM 的 prompt 示例、中文注释里的对齐。`skipStrings` 默认已开；这里再放过注释与模板字面量，
      //   避免把"展示公文版式的全角空格"当非法空白误删(删了会改 prompt/版式语义)。常规代码缩进仍禁全角空格。
      'no-irregular-whitespace': ['error', { skipComments: true, skipTemplates: true }],
    },
  },
  {
    files: ['**/*.test.ts', '**/*.spec.ts', '**/__tests__/**/*.ts'],
    rules: {
      'no-console': 'off',
    },
  },
  prettierConfig,
);
