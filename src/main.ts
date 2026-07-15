import { bootstrapApplication } from '@angular/platform-browser';
import { AppComponent } from './app/app.component';
import { provideHttpClient } from '@angular/common/http';
import { importProvidersFrom } from '@angular/core';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';

import { inject } from '@vercel/analytics';

inject();

bootstrapApplication(AppComponent, {
  providers: [
    provideHttpClient(),
    importProvidersFrom(FormsModule, ReactiveFormsModule)
    // Cualquier otra configuración de rutas o providers va directamente aquí
  ]
}).catch((err) => console.error(err));
