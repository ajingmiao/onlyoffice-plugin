import { logger } from '../core/logger.js';
import { EVENTS } from '../core/constants.js';

export class HostBridge {
  constructor(gateway = tryGetGateway()) {
    this.gateway = gateway;
  }

  sendInfo(command, data) {
    console.log('🏠 HostBridge.sendInfo 被调用:', { command, data });
    console.log('🔍 Gateway 状态:', {
      gateway: !!this.gateway,
      sendInfo: typeof this.gateway?.sendInfo,
      hasWindow: !!window,
      hasParent: !!window.parent,
      hasCommon: !!window.parent?.Common,
      hasGateway: !!window.parent?.Common?.Gateway
    });

    try {
      if (this.gateway?.sendInfo) {
        console.log('📤 正在调用 gateway.sendInfo...');
        this.gateway.sendInfo({ command, data });
        console.log('✅ gateway.sendInfo 调用成功');
      } else {
        console.warn('⚠️ gateway.sendInfo 不可用');
      }
    } catch (e) {
      console.error('❌ gateway.sendInfo 调用异常:', e);
    }
  }

  onInternalCommand(cb) {
    try {
      this.gateway?.on?.('internalcommand', (payload) => {
        cb?.(payload);
      });
    } catch (e) {
      logger.warn('Gateway.on unavailable');
    }
  }
}

function tryGetGateway() {
  try { return window.parent?.Common?.Gateway || null; } catch (e) { return null; }
}
