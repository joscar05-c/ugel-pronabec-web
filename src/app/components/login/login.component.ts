import { Component, ChangeDetectorRef, EventEmitter, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { SupabaseService } from '../../services/supabase.service';

@Component({
  selector: 'app-login',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './login.component.html',
})
export class LoginComponent {
  @Output() loginExitoso = new EventEmitter<void>();
  @Output() volver = new EventEmitter<void>();

  emailAdmin: string = '';
  passwordAdmin: string = '';
  errorLogin: string = '';

  constructor(
    private supabaseService: SupabaseService,
    private cdr: ChangeDetectorRef,
  ) {}

  onVolver() {
    this.volver.emit();
  }

  async verificarLogin() {
    if (!this.emailAdmin || !this.passwordAdmin) {
      this.errorLogin = 'Ingrese correo y contraseña.';
      return;
    }
    this.errorLogin = '';
    const { data, error } = await this.supabaseService.iniciarSesionAdmin(
      this.emailAdmin,
      this.passwordAdmin,
    );
    if (error) {
      this.errorLogin = 'Credenciales incorrectas.';
    } else {
      this.loginExitoso.emit();
    }
    this.cdr.detectChanges();
  }
}
