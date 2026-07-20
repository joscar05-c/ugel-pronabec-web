import { Component, ChangeDetectorRef, OnInit } from '@angular/core';
import { CommonModule, Location } from '@angular/common';
import { SupabaseService } from './services/supabase.service';
import { RegistroComponent } from './components/registro/registro.component';
import { LoginComponent } from './components/login/login.component';
import { AdminComponent } from './components/admin/admin.component';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, RegistroComponent, LoginComponent, AdminComponent],
  template: `
    <app-registro
      *ngIf="vistaActual === 'registro'"
      (navegarAAdmin)="cambiarVista('loginAdmin')">
    </app-registro>

    <app-login
      *ngIf="vistaActual === 'loginAdmin'"
      (loginExitoso)="onLoginExitoso()"
      (volver)="cambiarVista('registro')">
    </app-login>

    <app-admin
      *ngIf="vistaActual === 'admin'"
      (salir)="onSalir()">
    </app-admin>
  `,
})
export class AppComponent implements OnInit {
  vistaActual: string = 'registro';

  constructor(
    private supabaseService: SupabaseService,
    private cdr: ChangeDetectorRef,
    private location: Location,
  ) {}

  async ngOnInit() {
    await this.evaluarRutaActual();

    this.location.subscribe(async () => {
      await this.evaluarRutaActual();
    });
  }

  async evaluarRutaActual() {
    const { data } = await this.supabaseService.obtenerSesion();
    const rutaActual = this.location.path();

    if (data?.session) {
      if (rutaActual === '/admin') {
        this.vistaActual = 'admin';
      } else {
        this.vistaActual = 'registro';
      }
    } else {
      if (rutaActual === '/admin') {
        this.vistaActual = 'loginAdmin';
      } else {
        this.vistaActual = 'registro';
      }
    }

    this.cdr.detectChanges();
  }

  cambiarVista(nuevaVista: string) {
    this.vistaActual = nuevaVista;
    if (nuevaVista === 'registro') this.location.go('/');
    if (nuevaVista === 'loginAdmin') this.location.go('/admin');
    if (nuevaVista === 'admin') this.location.go('/admin');
    this.cdr.detectChanges();
  }

  async onLoginExitoso() {
    this.cambiarVista('admin');
  }

  async onSalir() {
    await this.supabaseService.cerrarSesion();
    this.cambiarVista('registro');
  }
}
