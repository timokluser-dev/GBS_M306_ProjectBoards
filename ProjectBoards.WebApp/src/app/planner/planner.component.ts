import { Component, OnInit } from "@angular/core";
import { HttpClient } from "@angular/common/http";
import { PlannerBuckets } from "../shared/interfaces/planner-buckets";
import { PlannerTasks } from "../shared/interfaces/planner-tasks";
import { ActivatedRoute, Router } from "@angular/router";
import { BucketChart } from "./bucket-chart";

@Component({
  selector: "app-planner",
  templateUrl: "./planner.component.html",
  styleUrls: ["./planner.component.css"],
})
export class PlannerComponent implements OnInit {
  private planId: string;
  private plannerBuckets: PlannerBuckets;
  public chartData: BucketChart[] = new Array();
  private bucketCount: number;

  constructor(
    private http: HttpClient,
    private route: ActivatedRoute,
    private router: Router
  ) {}

  ngOnInit() {
    this.route.params.subscribe((params) => {
      this.planId = params.planId;
      this.getChart();
    });
  }

  private getChart(): void {
    let plannerUri = `https://graph.microsoft.com/v1.0/planner/plans/${this.planId}/buckets`;
    this.http.get(plannerUri).subscribe({
      next: (_plannerBuckets) => {
        this.plannerBuckets = _plannerBuckets as PlannerBuckets;
        this.bucketCount = this.plannerBuckets.value.length;

        this.plannerBuckets.value.forEach((bucket) => {
          let bucketUri = `https://graph.microsoft.com/v1.0/planner/buckets/${bucket.id}/tasks`;
          this.http.get(bucketUri).subscribe({
            next: (_tasks) => {
              let bucketTasks = _tasks as PlannerTasks;

              let tasksCompleted: number = 0;
              let tasksCount: number = bucketTasks.value.length;

              bucketTasks.value.forEach(
                (task) =>
                  (tasksCompleted += this.percentToDecimal(
                    task.percentComplete
                  ))
              );

              let percentageCompleted: number =
                tasksCount != 0
                  ? Math.round((100 / tasksCount) * tasksCompleted)
                  : 0;

              let chart: BucketChart = {
                name: bucket.name,
                value: percentageCompleted,
                tasksInTotal: tasksCount,
                tasksCompleted: tasksCompleted,
              };

              this.addChart(chart);
            },
          });
        });
      },
      error: (_error) => {
        _error.status == 403 &&
          this.router.navigate(["/"]) &&
          console.warn(
            `Plan '${this.planId}' is not in the user-tenant or user has no access to it.`
          );
      },
    });
  }

  private percentToDecimal(percent: number): number {
    return percent / 100;
  }

  public isChartReady(): boolean {
    return this.chartData.length == this.bucketCount;
  }

  private addChart(data: BucketChart): void {
    this.chartData.push(data);
    this.sortCharts();
  }

  private sortCharts(): void {
    this.chartData = this.chartData.sort((a, b) =>
      a.name < b.name ? -1 : a.name > b.name ? 1 : 0
    );
  }
}
