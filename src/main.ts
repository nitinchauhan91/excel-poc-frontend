import 'zone.js';
import { bootstrapApplication } from '@angular/platform-browser';
import { appConfig } from './app/app.config';
import { AppComponent } from './app/app';

declare const Office: any;

Office.onReady(() => {
  bootstrapApplication(AppComponent, appConfig)
    .catch((err) => console.error(err));
});