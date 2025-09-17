export class EventBus {
    constructor({ hostBridge }) {
      this.host = hostBridge;
    }
    emit(command, data) {
      this.host.sendInfo(command, data);
    }
  }
  