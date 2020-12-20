import { Component, OnInit } from "@angular/core";
import { MsalService } from "@azure/msal-angular";
import { HttpClient } from "@angular/common/http";
import { GraphProfile } from "../shared/interfaces/graph-profile";
import { Router } from "@angular/router";

const GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0/me";

@Component({
  selector: "app-profile",
  templateUrl: "./profile.component.html",
  styleUrls: ["./profile.component.css"],
})
export class ProfileComponent implements OnInit {
  public profile: GraphProfile | null;

  constructor(
    private authService: MsalService,
    private http: HttpClient,
    private router: Router
  ) {}

  ngOnInit() {
    this.getProfile();
  }

  private getProfile() {
    this.http.get(GRAPH_ENDPOINT).subscribe({
      next: (profile) => {
        this.profile = profile as GraphProfile;
      },
      error: (_) => {
        this.router.navigate(["/"]);
      },
    });
  }
}
