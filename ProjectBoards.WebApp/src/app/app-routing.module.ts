import { NgModule } from "@angular/core";
import { Routes, RouterModule } from "@angular/router";
import { ProfileComponent } from "./profile/profile.component";
import { MsalGuard } from "@azure/msal-angular";
import { HomeComponent } from "./home/home.component";
import { PlannerComponent } from "./planner/planner.component";

const routes: Routes = [
  {
    path: "profile",
    component: ProfileComponent,
    canActivate: [MsalGuard],
  },
  {
    path: "planner/:planId",
    component: PlannerComponent,
    canActivate: [MsalGuard],
  },
  {
    path: "",
    component: HomeComponent,
  },
  {
    path: "**",
    redirectTo: "/",
  }
];

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: false })],
  exports: [RouterModule],
})
export class AppRoutingModule {}
