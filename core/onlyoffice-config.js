// core/onlyoffice-config.js - OnlyOffice统一配置管理
export class OnlyOfficeConfig {
    constructor() {
        // 默认配置
        this.defaultConfig = {
            // 服务器地址配置
            serverUrl: 'http://222.187.11.98:8918',
            apiPath: '/web-apps/apps/api/documents/api.js',
            editorPath: '/documenteditor/main/index.html',

            // 环境配置
            environment: 'production', // 'development' | 'production' | 'staging'

            // 调试配置
            debug: true,
            enableConsoleLog: true
        };

        // 环境特定配置
        this.environments = {
            development: {
                serverUrl: 'http://localhost:8080',
                debug: true,
                enableConsoleLog: true
            },
            staging: {
                serverUrl: 'http://222.187.11.98:8918',
                debug: true,
                enableConsoleLog: true
            },
            production: {
                serverUrl: 'http://222.187.11.98:8918',
                debug: false,
                enableConsoleLog: false
            }
        };

        this.currentConfig = this.mergeConfigs();
        this.logConfig();
    }

    // 合并配置
    mergeConfigs() {
        const envConfig = this.environments[this.defaultConfig.environment] || {};
        return {
            ...this.defaultConfig,
            ...envConfig
        };
    }

    // 获取完整的服务器URL
    getServerUrl() {
        return this.currentConfig.serverUrl;
    }

    // 获取API脚本URL
    getApiUrl() {
        return this.currentConfig.serverUrl + this.currentConfig.apiPath;
    }

    // 获取编辑器URL
    getEditorUrl() {
        return this.currentConfig.serverUrl + this.currentConfig.editorPath;
    }

    // 获取完整的OnlyOffice配置对象
    getOnlyOfficeConfig(options = {}) {
        const config = {
            type: options.type || 'desktop',
            width: options.width || '100%',
            height: options.height || '600px',
            documentType: options.documentType || 'word',

            // 重要：确保这些URL是完整的
            documentServerUrl: this.getServerUrl() + '/',
            server: this.getServerUrl(),
            apiUrl: this.getApiUrl(),

            document: {
                fileType: options.fileType || 'docx',
                key: options.key || this.generateKey(),
                title: options.title || 'OnlyOffice文档',
                url: options.documentUrl,
                permissions: {
                    comment: true,
                    copy: true,
                    download: true,
                    edit: true,
                    fillForms: true,
                    modifyContentControl: true,
                    modifyFilter: true,
                    print: true,
                    protect: false,
                    review: true
                }
            },

            editorConfig: {
                mode: options.mode || 'edit',
                lang: options.lang || 'zh-CN',
                user: {
                    id: options.userId || 'demo-user',
                    name: options.userName || '演示用户',
                    group: 'demo'
                },
                customization: {
                    autosave: false,
                    forcesave: false,
                    chat: false,
                    compactToolbar: false,
                    compactHeader: false,
                    feedback: { visible: false },
                    help: false,
                    hideRightMenu: false,
                    hideRulers: false,
                    integrationMode: 'embed',
                    macros: true,
                    macrosMode: 'enable',
                    mentionShare: false,
                    mobileForceView: false,
                    plugins: true,
                    reviewDisplay: 'original',
                    spellcheck: true,
                    submitForm: false,
                    toolbarHideFileName: false,
                    toolbarNoTabs: false,
                    trackChanges: false,
                    unit: 'cm',
                    zoom: 100,
                    goback: false,
                    about: false
                },
                events: options.events || {}
            }
        };

        if (this.currentConfig.debug) {
            this.log('生成OnlyOffice配置:', config);
        }

        return config;
    }

    // 生成文档密钥
    generateKey() {
        const timestamp = new Date().getTime();
        const random = Math.random().toString(36).substr(2, 9);
        return `doc_${timestamp}_${random}`;
    }

    // 切换环境
    switchEnvironment(env) {
        if (this.environments[env]) {
            this.defaultConfig.environment = env;
            this.currentConfig = this.mergeConfigs();
            this.log(`已切换到环境: ${env}`, this.currentConfig);
            return true;
        }
        this.log(`未知环境: ${env}`);
        return false;
    }

    // 更新配置
    updateConfig(updates) {
        this.currentConfig = { ...this.currentConfig, ...updates };
        this.log('配置已更新:', this.currentConfig);
    }

    // 获取当前配置
    getCurrentConfig() {
        return { ...this.currentConfig };
    }

    // 日志输出
    log(message, data = null) {
        if (this.currentConfig.enableConsoleLog) {
            console.log(`[OnlyOfficeConfig] ${message}`, data || '');
        }
    }

    // 配置初始化日志
    logConfig() {
        if (this.currentConfig.debug) {
            this.log('OnlyOffice配置初始化完成');
            this.log('当前环境:', this.currentConfig.environment);
            this.log('服务器地址:', this.getServerUrl());
            this.log('API地址:', this.getApiUrl());
            this.log('编辑器地址:', this.getEditorUrl());
        }
    }

    // 验证配置
    validateConfig() {
        const errors = [];

        if (!this.currentConfig.serverUrl) {
            errors.push('服务器地址未配置');
        }

        try {
            new URL(this.getServerUrl());
        } catch (e) {
            errors.push('服务器地址格式无效');
        }

        if (errors.length > 0) {
            this.log('配置验证失败:', errors);
            return { valid: false, errors };
        }

        this.log('配置验证通过');
        return { valid: true, errors: [] };
    }

    // 测试连接
    async testConnection() {
        try {
            this.log('测试OnlyOffice服务器连接...');
            const response = await fetch(this.getApiUrl(), {
                method: 'HEAD',
                mode: 'no-cors' // 避免CORS问题
            });

            this.log('连接测试成功');
            return { success: true, message: '服务器连接正常' };
        } catch (error) {
            this.log('连接测试失败:', error.message);
            return { success: false, message: `连接失败: ${error.message}` };
        }
    }
}

// 创建全局单例
export const onlyOfficeConfig = new OnlyOfficeConfig();

// 导出便捷方法
export const getOnlyOfficeConfig = (options) => onlyOfficeConfig.getOnlyOfficeConfig(options);
export const getServerUrl = () => onlyOfficeConfig.getServerUrl();
export const getApiUrl = () => onlyOfficeConfig.getApiUrl();
export const switchEnvironment = (env) => onlyOfficeConfig.switchEnvironment(env);