module.exports = {
    extends: ['@commitlint/config-conventional'],
    rules: {
        'type-enum': [
            2,
            'always',
            ['feat', 'fix', 'hotfix', 'docs', 'style', 'refactor', 'test', 'build', 'ci', 'chore'],
        ],
        'scope-case': [0, 'always', 'camel-case'],
        'scope-empty': [2, 'never'],
    },
    prompt: {
        questions: {
            type: {
                enum: {
                    hotfix: {
                        description: 'A online bug fix',
                        title: 'Hotfix',
                        emoji: '🔥',
                    },
                },
            },
        },
    },
};
