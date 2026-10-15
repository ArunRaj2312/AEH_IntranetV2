import React, { useEffect, useState } from "react";
import type { SPFI } from "@pnp/sp";
import "./QuickLinks.css";

export interface IQuickLink {
  Title: string;
  Description: string;
  Image: string;
  Link: { Url: string; Description?: string } | string;
  Color: string;
}

interface IQuickLinksProps {
  /** Configured PnPjs SPFI instance for the current SharePoint site. */
  sp: SPFI;
  context: any;
  /** Optional list title; defaults to the source web part's list. */
  listTitle?: string;
}

function getImageUrl(imageValue: string, origin: string): string {
  if (!imageValue) return "";
  try {
    const parsed = JSON.parse(imageValue);
    const relativeUrl = parsed?.serverRelativeUrl;
    if (relativeUrl) return `${origin}${relativeUrl}`;
  } catch {
    // Allow an image field that already contains a URL.
  }
  return imageValue;
}

function getLinkUrl(link: IQuickLink["Link"]): string {
  if (typeof link === "string") return link;
  return link?.Url ?? "";
}

const MainComponent: React.FC<IQuickLinksProps> = ({
  sp,
  context,
  listTitle = "IntranetComponents",
}) => {
  const [links, setLinks] = useState<IQuickLink[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    let active = true;

    const loadLinks = async () => {
      try {
        setLoading(true);
        setError("");
        const items: any = await sp.web.lists
          .getByTitle(listTitle)
          .items.filter("WebParts eq 'QuickLinks'")
          .select("Title", "Description", "Image", "Link", "Color", "ID")();

        let mapped: IQuickLink[] = [];
        items.forEach((item: any) => {
          let objImage = JSON.parse(item.Image);

          let imgUrl: string =
            context?.pageContext?.site?.absoluteUrl +
            "/Lists/IntranetComponents" +
            "/Attachments/" +
            item.ID +
            "/" +
            objImage.fileName;

          mapped.push({
            Title: item.Title ?? "",
            Description: item.Description ?? "",
            Image: imgUrl ?? "",
            Link: item.Link ?? "",
            Color: item.Color || "#0098bd",
          });
        });

        if (active) setLinks(mapped);
      } catch (e) {
        console.error("Unable to load Quick Links from SharePoint.", e);
        if (active)
          setError("Quick Links could not be loaded. Please try again later.");
      } finally {
        if (active) setLoading(false);
      }
    };

    void loadLinks();
    return () => {
      active = false;
    };
  }, [sp, listTitle]);

  if (loading)
    return <div className="quickLinksMessage">Loading quick links…</div>;
  if (error)
    return (
      <div className="quickLinksMessage quickLinksError" role="alert">
        {error}
      </div>
    );

  return (
    <section className="quickLinks" aria-label="Quick links">
      <div className="quickLinksGrid">
        {links.map((link, index) => {
          const href = getLinkUrl(link.Link);
          const imageUrl = getImageUrl(link.Image, window.location.origin);

          return (
            <article
              className="quickLinkCard"
              style={{ backgroundColor: link.Color }}
              key={`${link.Title}-${href}-${index}`}
            >
              <a
                className="quickLinkAnchor"
                href={href || undefined}
                target="_blank"
                rel="noopener noreferrer"
                aria-label={`${link.Title}${link.Description ? `: ${link.Description}` : ""}`}
              >
                <div className="quickLinkImageSection">
                  {imageUrl && (
                    <img className="quickLinkImage" src={imageUrl} alt="" />
                  )}
                </div>
                <div className="quickLinkContent">
                  <h2>{link.Title}</h2>
                  <p>{link.Description}</p>
                </div>
              </a>
            </article>
          );
        })}
      </div>
      {!links.length && (
        <div className="quickLinksMessage">No quick links found.</div>
      )}
    </section>
  );
};

export default MainComponent;
