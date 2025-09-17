import { createContainer } from './app/container.js';
import { bootstrap } from './app/lifecycle.js';

const di = createContainer();
bootstrap(di);
