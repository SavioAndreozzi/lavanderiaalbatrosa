CREATE TABLE "laundry" (
	"id" integer PRIMARY KEY,
	"records" jsonb DEFAULT '[]' NOT NULL,
	"version" integer DEFAULT 0 NOT NULL,
	"updated_at" timestamp with time zone DEFAULT now() NOT NULL
);
