import { useState, useMemo, useEffect, useCallback, useRef } from "react";
import { MsalProvider, useMsal, useIsAuthenticated } from "@azure/msal-react";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend } from "recharts";
import { msalInstance, loginRequest } from "./authConfig";
import { fetchLivePilotData } from "./graphService";

const AUTO_REFRESH_MS = 60 * 1000; // 1 minute

// ─── FALLBACK SAMPLE DATA (used only if live fetch hasn't loaded / fails) ──
const FALLBACK_ONBOARDING = [{"txn_id":"06DDC8D8030A4D54","authorize":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"rKiUbXNQ4Ar2SrKv1fCcr5PpNqGxwp6Zvy","created_at":"8/6/2026, 4:00:06 pm","user_name":"Sophatnith Khorn","agent_name":null},{"txn_id":"3B9ABDB2C052A6D3","authorize":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"r3ALnmoNXAXHaTKFHHRasdxco1UvuimDpF","created_at":"8/6/2026, 3:57:42 pm","user_name":"Virak Rim","agent_name":null},{"txn_id":"40A4E20B210D07D9","authorize":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","destination":"rpAqLirhRZGcF4u5uCWS5uvxxyjkA6kLiT","created_at":"8/6/2026, 3:55:43 pm","user_name":"Virak Rim","agent_name":"Virak Rim"},{"txn_id":"D0610B336BF2608F","authorize":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"r4bxqBDA1DePz9dsfxjHTn6sgkaRadUpLX","created_at":"8/6/2026, 2:42:56 pm","user_name":"Laga Fatima","agent_name":null},{"txn_id":"0F4C0D36427866C2","authorize":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"rwuSWJm1x4zWTg8z9sxCEAt9M7CfjPQp3d","created_at":"8/6/2026, 2:07:19 pm","user_name":"Sophatnith Khorn","agent_name":null},{"txn_id":"B012A697A9F120E9","authorize":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"rwm3ZBXXLm7JN5u7ffQBGzGe7yyuGodRYw","created_at":"8/6/2026, 2:05:23 pm","user_name":"Sophatnith Khoen","agent_name":null},{"txn_id":"2FA87C1540DD6C50","authorize":"rhBZ3Pj6LXXpEZqaSEca8at8cbkWPKnvWU","destination":"rfv1wKTZQEsWoi1Ku5YVdgTWrRzCmFg9qo","created_at":"8/6/2026, 2:03:31 pm","user_name":"Sophia Kuo","agent_name":"SOP Phnith"},{"txn_id":"A5FF2D6E817C3B1A","authorize":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"rwuSWJm1x4zWTg8z9sxCEAt9M7CfjPQp3d","created_at":"8/6/2026, 2:07:19 pm","user_name":"Gugujg Tugug","agent_name":null},{"txn_id":"C1234567890ABCDE","authorize":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","destination":"rwm3ZBXXLm7JN5u7ffQBGzGe7yyuGodRYw","created_at":"8/6/2026, 1:27:14 pm","user_name":"Djjh Dff","agent_name":"Virak Rim"},{"txn_id":"D1234567890ABCDE","authorize":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","destination":"r4bxqBDA1DePz9dsfxjHTn6sgkaRadUpLX","created_at":"8/6/2026, 1:24:46 pm","user_name":"Sophia Keo","agent_name":"Virak Rim"},{"txn_id":"E1234567890ABCDE","authorize":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","destination":"r4bxqBDA1DePz9dsfxjHTn6sgkaRadUpLX","created_at":"8/6/2026, 1:04:12 pm","user_name":"Virak Rin","agent_name":"Virak Rim"},{"txn_id":"F1234567890ABCDE","authorize":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","destination":"rKRR2yCwk8bgW1pPom4LPD4doa3Fd58wyD","created_at":"8/6/2026, 12:14:50 pm","user_name":"Banana Saka","agent_name":"Virak Rim"},{"txn_id":"G1234567890ABCDE","authorize":"rhBZ3Pj6LXXpEZqaSEca8at8cbkWPKnvWU","destination":"rKRR2yCwk8bgW1pPom4LPD4doa3Fd58wyD","created_at":"8/6/2026, 12:08:32 pm","user_name":"SOPHEAK KIM","agent_name":"SOP Phnith"},{"txn_id":"H1234567890ABCDE","authorize":"rhBZ3Pj6LXXpEZqaSEca8at8cbkWPKnvWU","destination":"rMcePk96JbpTNJtR36HiixoyTayj4XVXy3","created_at":"8/6/2026, 10:02:05 am","user_name":"Sophatnith Khorn","agent_name":"SOP Phnith"},{"txn_id":"I1234567890ABCDE","authorize":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"r3pJLjn6Te6wi4guD7Cc8EynadV8RUxWPj","created_at":"8/6/2026, 7:33:25 am","user_name":"Sophatnith Khon","agent_name":null},{"txn_id":"J1234567890ABCDE","authorize":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"rGPX5yQoFgVm8bKaFsTMyTGiL3eYnG3uFE","created_at":"8/6/2026, 7:25:49 am","user_name":"Sophia Kt","agent_name":null},{"txn_id":"K1234567890ABCDE","authorize":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"rNZxBbsdjejur8eAsJEm4DK9MKAjwuXzVc","created_at":"8/6/2026, 7:23:46 am","user_name":"Sophia Gvfh","agent_name":null},{"txn_id":"L1234567890ABCDE","authorize":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"rNF16gjqAaWnUfLNgx8ZFtzPoUsC3f6caY","created_at":"8/6/2026, 7:22:07 am","user_name":"Spgcv Gghfc","agent_name":null},{"txn_id":"M1234567890ABCDE","authorize":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"rBSbbcAiv2tQhkw7njztAfkTAmQpWHTVfZ","created_at":"8/6/2026, 7:09:09 am","user_name":"Sophia Kaoo","agent_name":null},{"txn_id":"N1234567890ABCDE","authorize":"raFBRENgtjTXDFTH16E3JtiSJtxLFW3WWa","destination":"rnCdwMG6BagQJ6VjqiYrYpoUNXAdwKfcQ8","created_at":"4/6/2026, 5:31:13 pm","user_name":"Sophia Httt","agent_name":"La Fatima"},{"txn_id":"O1234567890ABCDE","authorize":"raFBRENgtjTXDFTH16E3JtiSJtxLFW3WWa","destination":"rHCY8hL7wrRhbAv2rPaRfq3Tb5HmdSDgZH","created_at":"4/6/2026, 5:22:04 pm","user_name":"Sopheak Kim","agent_name":"La Fatima"},{"txn_id":"P1234567890ABCDE","authorize":"raFBRENgtjTXDFTH16E3JtiSJtxLFW3WWa","destination":"rhbWbnXEaod2gJFuSsJ1tgv3Z6stX56zrK","created_at":"4/6/2026, 5:17:22 pm","user_name":"Virak Rin","agent_name":"La Fatima"},{"txn_id":"Q1234567890ABCDE","authorize":"raFBRENgtjTXDFTH16E3JtiSJtxLFW3WWa","destination":"rrhc4wJ1bJiyoFGujTJExo9TV6egc5NWTN","created_at":"4/6/2026, 3:53:53 pm","user_name":"Sopheak Kim","agent_name":"La Fatima"},{"txn_id":"R1234567890ABCDE","authorize":"raFBRENgtjTXDFTH16E3JtiSJtxLFW3WWa","destination":"r36HwhgPSzicQpUiFaimz9nMPz72Py5Vx5","created_at":"4/6/2026, 12:32:09 pm","user_name":"Sophia Keo","agent_name":"La Fatima"},{"txn_id":"S1234567890ABCDE","authorize":"raFBRENgtjTXDFTH16E3JtiSJtxLFW3WWa","destination":"r9GEtg55NSPe643WJgyqb6QjautyA6fB6k","created_at":"4/6/2026, 12:13:20 pm","user_name":"Ta Zayia","agent_name":"La Fatima"},{"txn_id":"T1234567890ABCDE","authorize":"raFBRENgtjTXDFTH16E3JtiSJtxLFW3WWa","destination":"rKRMBgGAUdNvTNAoQiDqivGocWVYQF2KZH","created_at":"4/6/2026, 11:47:25 am","user_name":"Ha Ja","agent_name":"La Fatima"},{"txn_id":"U1234567890ABCDE","authorize":"rGyuTbbiNwGhuoLDteXiP3gBtVTx5uTvu","destination":"rNcUQijo2K8pJdvhsMTiLVoviNcgGkf1pj","created_at":"27/5/2026, 1:54:12 pm","user_name":"BUNTHOEUN PECH","agent_name":"BUNTHOEUN PECH"},{"txn_id":"V1234567890ABCDE","authorize":"rKXnG3cu6Zaoqc7Tm8qWYngpy4bssFswgQ","destination":"rKXnG3cu6Zaoqc7Tm8qWYngpy4bssFswgQ","created_at":"27/5/2026, 9:15:09 am","user_name":"Alpha Beta","agent_name":"Brian Moeng"},{"txn_id":"W1234567890ABCDE","authorize":"rKXnG3cu6Zaoqc7Tm8qWYngpy4bssFswgQ","destination":"rUQvGTqrizAJXKHavPJ8Q7994dyc8U5e5z","created_at":"27/5/2026, 8:55:17 am","user_name":"Gala Lifa","agent_name":"Brian Moeng"},{"txn_id":"X1234567890ABCDE","authorize":"rKXnG3cu6Zaoqc7Tm8qWYngpy4bssFswgQ","destination":"rP2avL5s4FGAreYQNdqXN8sTXgzqwEVGDh","created_at":"27/5/2026, 8:10:22 am","user_name":"Sophatnith Nith","agent_name":"Brian Moeng"},{"txn_id":"Y1234567890ABCDE","authorize":"rgCurhaXKzM67cybcTKnzDLrcjeM3usRo","destination":"rH1xUvjGHEDpbin1cVwpb25AjAiPnsio4Y","created_at":"3/6/2026, 6:04:55 pm","user_name":"Fa Ebola","agent_name":"CINDY SEAN SIBANDA"},{"txn_id":"Z1234567890ABCDE","authorize":"rgCurhaXKzM67cybcTKnzDLrcjeM3usRo","destination":"rJGcvwqEVdHiK1CX6SFKVGN9dCQY1xLWvb","created_at":"3/6/2026, 6:04:46 pm","user_name":"Lo Po","agent_name":"CINDY SEAN SIBANDA"},{"txn_id":"AA234567890ABCDE","authorize":"raWsJizhUmqisuD38TNpptKA6jDDcdRR7X","destination":"rn3n1iGmPVt96nh8bnTXixQEdTxg9ddz6e","created_at":"4/6/2026, 5:50:57 am","user_name":"Soknov 8","agent_name":"Phone Agent"},{"txn_id":"BB234567890ABCDE","authorize":"raWsJizhUmqisuD38TNpptKA6jDDcdRR7X","destination":"rpzkgg2bbnkZvpQNSc947kMXgHNK239Wcc","created_at":"4/6/2026, 5:50:51 am","user_name":"Soknov 7","agent_name":"Phone Agent"},{"txn_id":"CC234567890ABCDE","authorize":"rWNsjHTPQbBkeC8QDyG57jSQ3qyqWPM94","destination":"rWNsjHTPQbBkeC8QDyG57jSQ3qyqWPM94","created_at":"3/6/2026, 1:12:59 pm","user_name":"Virak Rim","agent_name":"Soknov KP03"},{"txn_id":"DD234567890ABCDE","authorize":"rsvCoSB8PfpArJuKPyoAMBNTinxZWo5DhS","destination":"rfzpW1pibdwHrMZZJzC5MkGGeaWmMDqebK","created_at":"2/6/2026, 12:54:51 pm","user_name":"Khalifa Autuma","agent_name":"Sopheak Kim"},{"txn_id":"EE234567890ABCDE","authorize":"rsvCoSB8PfpArJuKPyoAMBNTinxZWo5DhS","destination":"r4nqC5bTZfSFjRZJEw7DeLYW7eifdAo4AV","created_at":"2/6/2026, 12:54:45 pm","user_name":"Fatima Autuma","agent_name":"Sopheak Kim"},{"txn_id":"FF234567890ABCDE","authorize":"rsvCoSB8PfpArJuKPyoAMBNTinxZWo5DhS","destination":"rfipyYwHavtVhHMiwoRdL9nq7hA8EAaLjM","created_at":"2/6/2026, 12:47:37 pm","user_name":"Sopheak Soundbox","agent_name":"Sopheak Kim"},{"txn_id":"GG234567890ABCDE","authorize":"rsvCoSB8PfpArJuKPyoAMBNTinxZWo5DhS","destination":"rPNGHMH7wFTN1j9X6rSFH3WCnmUJ5Wh46w","created_at":"2/6/2026, 7:17:58 am","user_name":"Chatha Neang","agent_name":"Sopheak Kim"},{"txn_id":"HH234567890ABCDE","authorize":"rsvCoSB8PfpArJuKPyoAMBNTinxZWo5DhS","destination":"rKFABhtrPLTTMP6tgsDge42CU6MtEW6CiF","created_at":"2/6/2026, 3:17:54 pm","user_name":"Sophia Kim","agent_name":"Sopheak Kim"},{"txn_id":"II234567890ABCDE","authorize":"rDfCe5i2GemgRpLNVX6m4CFNhvtgvqz8Dc","destination":"rSKddRvhZfwSs8mKJFMGKGor3qnCQ61Qf","created_at":"2/6/2026, 11:24:55 pm","user_name":"Zala Fanipa","agent_name":"VIRAK RIM"},{"txn_id":"JJ234567890ABCDE","authorize":"rDfCe5i2GemgRpLNVX6m4CFNhvtgvqz8Dc","destination":"rDfCe5i2GemgRpLNVX6m4CFNhvtgvqz8Dc","created_at":"2/6/2026, 11:34:59 pm","user_name":"Buyia Rajia","agent_name":"VIRAK RIM"},{"txn_id":"KK234567890ABCDE","authorize":"rUz2AoJVpLp2s66VbTM4qk6XPyEkRsnTfF","destination":"rUz2AoJVpLp2s66VbTM4qk6XPyEkRsnTfF","created_at":"29/5/2026, 6:54:39 am","user_name":"Kim Sopheak","agent_name":"Ya Yia"},{"txn_id":"LL234567890ABCDE","authorize":"rUz2AoJVpLp2s66VbTM4qk6XPyEkRsnTfF","destination":"r4ccx1vVv114ERb2busGshTSmczkk7JyHm","created_at":"28/5/2026, 3:58:53 pm","user_name":"VIRAK RIM","agent_name":"Ya Yia"},{"txn_id":"MM234567890ABCDE","authorize":"rU6LFdty5fj8M48vpEtfPDkj9tVkTGi5pF","destination":"r4f8V8CaTooiLN5P368LrdMxG1ZYvXXWPY","created_at":"25/5/2026, 11:51:51 am","user_name":"Minea CHHEM","agent_name":"Kim Sopheak"},{"txn_id":"NN234567890ABCDE","authorize":"rU6LFdty5fj8M48vpEtfPDkj9tVkTGi5pF","destination":"rnYTpwdG18w48eqTCcnKy2roHTJqsebDC7","created_at":"1/6/2026, 9:45:52 am","user_name":"Ki Lo","agent_name":"Kim Sopheak"},{"txn_id":"OO234567890ABCDE","authorize":"rpCYA1dHbQYrJwrRoLvSAEeoZL4sUD75gD","destination":"rpCYA1dHbQYrJwrRoLvSAEeoZL4sUD75gD","created_at":"4/6/2026, 9:25:12 am","user_name":"Alpha Been","agent_name":"Fa Ebola"},{"txn_id":"PP234567890ABCDE","authorize":"r9sJyRMFGuRvA3zNJ9iMgPuHyJSrjGb8Q2","destination":"r9sJyRMFGuRvA3zNJ9iMgPuHyJSrjGb8Q2","created_at":"29/5/2026, 7:28:16 am","user_name":"Virak Rim","agent_name":"Ebola Buyia"},{"txn_id":"QQ234567890ABCDE","authorize":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","destination":"rHeFyBLfapzGqNTj3JruyrygHJvqDKJ6em","created_at":"8/6/2026, 1:11:57 pm","user_name":"Sophia Keo","agent_name":"Virak Rim"},{"txn_id":"RR234567890ABCDE","authorize":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","destination":"rQhhUWvJbskfogSLtxFFV3gQgpr6fQjZ2x","created_at":"8/6/2026, 12:14:50 pm","user_name":"Sadi Mckenzie","agent_name":"Virak Rim"},{"txn_id":"SS234567890ABCDE","authorize":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","destination":"rUdkVUWAmN2j5v27cS7AYmRMvnJihonV2K","created_at":"5/6/2026, 4:29:22 pm","user_name":"ABC Shop","agent_name":"Virak Rim"},{"txn_id":"TT234567890ABCDE","authorize":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","destination":"rNbKX8wqqXKN4j1gr5djqFo7izYwJ56YZb","created_at":"4/6/2026, 8:51:05 am","user_name":"Spotify Khim","agent_name":"Virak Rim"},{"txn_id":"UU234567890ABCDE","authorize":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","destination":"rGx9zsMXZx9nMw4u3E6r2audasfuy4Ufo7","created_at":"4/6/2026, 8:58:02 am","user_name":"Rim Virak","agent_name":"Virak Rim"},{"txn_id":"VV234567890ABCDE","authorize":"rhBZ3Pj6LXXpEZqaSEca8at8cbkWPKnvWU","destination":"rLmt4wcm7FAJYCynBHs2HNjyYxt2i9uwcb","created_at":"4/6/2026, 8:45:14 am","user_name":"Sophatnith Khhorn","agent_name":"SOP Phnith"},{"txn_id":"WW234567890ABCDE","authorize":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","destination":"rLFBvQYs25jnVG44F8h2ZadxTR1MtGQ6sv","created_at":"3/6/2026, 3:57:29 pm","user_name":"VIRAK Rim","agent_name":"VIRAK Rim"},{"txn_id":"XX234567890ABCDE","authorize":"rMBhr9weybfq1CHzCYd49cMGkA5hjTdXqW","destination":"r3djCXTkUgJki3yfWDkYCQpD7YQ19dR4JB","created_at":"26/5/2026, 3:58:05 pm","user_name":"VIRAK RIM","agent_name":null}];

const FALLBACK_PAYMENTS = [{"txn_id":"9C0CB133FC4A113C","account":"rKiUbXNQ4Ar2SrKv1fCcr5PpNqGxwp6Zvy","destination":"rgC3cPzc6y2ky4knhAPGHmQVKQ4zLxghG","amount":1,"created_at":"8/6/2026, 4:43:52 pm","sender":"VIRAK RIM","receiver":"VIRAK RIM","included":true},{"txn_id":"B216AE834586CAE5","account":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"rKiUbXNQ4Ar2SrKv1fCcr5PpNqGxwp6Zvy","amount":100,"created_at":"8/6/2026, 4:00:17 pm","sender":"Sophia Gvfh","receiver":"VIRAK RIM","included":true},{"txn_id":"FB6EB5C56E1BBB35","account":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"r3ALnmoNXAXHaTKFHHRasdxco1UvuimDpF","amount":100,"created_at":"8/6/2026, 3:57:53 pm","sender":"Sophia Gvfh","receiver":"","included":true},{"txn_id":"906ACA860977C1D9","account":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","destination":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","amount":158,"created_at":"8/6/2026, 3:56:31 pm","sender":"Pich Sovanna","receiver":"Sophia Kaoo","included":true},{"txn_id":"6E05919DEB177833","account":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","destination":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","amount":128,"created_at":"8/6/2026, 3:56:01 pm","sender":"Sophia Kaoo","receiver":"Pich Sovanna","included":true},{"txn_id":"E17783A098AA9A98","account":"r3m2ouQkHp6RkYF5ZTQAEfn2fKDf2MQXHA","destination":"rpAqLirhRZGcF4u5uCWS5uvxxyjkA6kLiT","amount":100,"created_at":"8/6/2026, 3:55:52 pm","sender":"Sophia Gvfh","receiver":"VIRAK RIM","included":true},{"txn_id":"4044153173B3D53C","account":"rhFTq8XcfxbyqurtsEAVgn9QUSeks6zpBX","destination":"r4bxqBDA1DePz9dsfxjHTn6sgkaRadUpLX","amount":278,"created_at":"8/6/2026, 3:21:23 pm","sender":"VIRAK RIM","receiver":"Spgcv Gghfc","included":true},{"txn_id":"764D6EDE8C72637E","account":"rhFTq8XcfxbyqurtsEAVgn9QUSeks6zpBX","destination":"r4bxqBDA1DePz9dsfxjHTn6sgkaRadUpLX","amount":3333,"created_at":"8/6/2026, 3:21:00 pm","sender":"VIRAK RIM","receiver":"Spgcv Gghfc","included":true},{"txn_id":"A1111111111111AA","account":"rhFTq8XcfxbyqurtsEAVgn9QUSeks6zpBX","destination":"r4bxqBDA1DePz9dsfxjHTn6sgkaRadUpLX","amount":500,"created_at":"8/6/2026, 2:45:00 pm","sender":"VIRAK RIM","receiver":"Spgcv Gghfc","included":true},{"txn_id":"A2222222222222AA","account":"rhFTq8XcfxbyqurtsEAVgn9QUSeks6zpBX","destination":"r4bxqBDA1DePz9dsfxjHTn6sgkaRadUpLX","amount":800,"created_at":"8/6/2026, 2:30:00 pm","sender":"VIRAK RIM","receiver":"Laga Fatima","included":true},{"txn_id":"A3333333333333AA","account":"rhFTq8XcfxbyqurtsEAVgn9QUSeks6zpBX","destination":"r4bxqBDA1DePz9dsfxjHTn6sgkaRadUpLX","amount":1200,"created_at":"8/6/2026, 2:15:00 pm","sender":"VIRAK RIM","receiver":"Laga Fatima","included":true},{"txn_id":"A4444444444444AA","account":"rhFTq8XcfxbyqurtsEAVgn9QUSeks6zpBX","destination":"r4bxqBDA1DePz9dsfxjHTn6sgkaRadUpLX","amount":650,"created_at":"8/6/2026, 1:55:00 pm","sender":"VIRAK RIM","receiver":"Sophia Gvfh","included":true},{"txn_id":"B1111111111111BB","account":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","destination":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","amount":444,"created_at":"8/6/2026, 1:20:00 pm","sender":"La Fatima","receiver":"Sophia Kaoo","included":true},{"txn_id":"B2222222222222BB","account":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","destination":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","amount":900,"created_at":"8/6/2026, 12:50:00 pm","sender":"La Fatima","receiver":"Pich Sovanna","included":true},{"txn_id":"B3333333333333BB","account":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","destination":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","amount":2000,"created_at":"8/6/2026, 12:10:00 pm","sender":"La Fatima","receiver":"Pich Sovanna","included":true},{"txn_id":"B4444444444444BB","account":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","destination":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","amount":500,"created_at":"8/6/2026, 11:40:00 am","sender":"La Fatima","receiver":"Sophia Kaoo","included":true},{"txn_id":"B5555555555555BB","account":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","destination":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","amount":750,"created_at":"8/6/2026, 11:00:00 am","sender":"La Fatima","receiver":"Sophia Gvfh","included":true},{"txn_id":"C1111111111111CC","account":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","destination":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","amount":333,"created_at":"7/6/2026, 4:30:00 pm","sender":"Sophia Kaoo","receiver":"La Fatima","included":true},{"txn_id":"C2222222222222CC","account":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","destination":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","amount":250,"created_at":"7/6/2026, 3:45:00 pm","sender":"Sophia Kaoo","receiver":"Pich Sovanna","included":true},{"txn_id":"C3333333333333CC","account":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","destination":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","amount":198,"created_at":"7/6/2026, 2:20:00 pm","sender":"Spgcv Gghfc","receiver":"VIRAK RIM","included":true},{"txn_id":"C4444444444444CC","account":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","destination":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","amount":681,"created_at":"7/6/2026, 1:50:00 pm","sender":"Spgcv Gghfc","receiver":"La Fatima","included":true},{"txn_id":"C5555555555555CC","account":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","destination":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","amount":125,"created_at":"7/6/2026, 12:30:00 pm","sender":"Spgcv Gghfc","receiver":"La Fatima","included":true},{"txn_id":"D1111111111111DD","account":"rKiUbXNQ4Ar2SrKv1fCcr5PpNqGxwp6Zvy","destination":"rgC3cPzc6y2ky4knhAPGHmQVKQ4zLxghG","amount":5000,"created_at":"7/6/2026, 10:00:00 am","sender":"VIRAK RIM","receiver":"Sophia Kaoo","included":true},{"txn_id":"D2222222222222DD","account":"rKiUbXNQ4Ar2SrKv1fCcr5PpNqGxwp6Zvy","destination":"rgC3cPzc6y2ky4knhAPGHmQVKQ4zLxghG","amount":9800,"created_at":"6/6/2026, 4:00:00 pm","sender":"VIRAK RIM","receiver":"Pich Sovanna","included":true},{"txn_id":"D3333333333333DD","account":"rKiUbXNQ4Ar2SrKv1fCcr5PpNqGxwp6Zvy","destination":"rgC3cPzc6y2ky4knhAPGHmQVKQ4zLxghG","amount":1000,"created_at":"6/6/2026, 2:30:00 pm","sender":"VIRAK RIM","receiver":"Spgcv Gghfc","included":true},{"txn_id":"D4444444444444DD","account":"rKiUbXNQ4Ar2SrKv1fCcr5PpNqGxwp6Zvy","destination":"rgC3cPzc6y2ky4knhAPGHmQVKQ4zLxghG","amount":700,"created_at":"6/6/2026, 11:00:00 am","sender":"VIRAK RIM","receiver":"La Fatima","included":true},{"txn_id":"E1111111111111EE","account":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","destination":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","amount":658,"created_at":"6/6/2026, 9:00:00 am","sender":"La Fatima","receiver":"Sophia Kaoo","included":true},{"txn_id":"E2222222222222EE","account":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","destination":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","amount":870,"created_at":"5/6/2026, 5:00:00 pm","sender":"La Fatima","receiver":"Sophia Gvfh","included":true},{"txn_id":"E3333333333333EE","account":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","destination":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","amount":798,"created_at":"5/6/2026, 3:00:00 pm","sender":"La Fatima","receiver":"Spgcv Gghfc","included":true},{"txn_id":"E4444444444444EE","account":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","destination":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","amount":618,"created_at":"5/6/2026, 1:00:00 pm","sender":"Pich Sovanna","receiver":"La Fatima","included":true},{"txn_id":"E5555555555555EE","account":"rfaR9JTLbafEvZWsczz6NRBbugDNgj5i7Z","destination":"rBB93dTjgFbgtcNRk6AEExs1duLne5YZPA","amount":281,"created_at":"5/6/2026, 10:00:00 am","sender":"Pich Sovanna","receiver":"La Fatima","included":true}];

// ─── INCENTIVE CALCULATION ENGINE ───────────────────────────────────────────
function calcMerchantIncentives(merchantName, onboardingData, paymentData) {
  const includedPayments = paymentData.filter(p => p.included && (p.sender === merchantName || p.receiver === merchantName));
  const sentPayments = paymentData.filter(p => p.included && p.sender === merchantName);
  const txnCount = includedPayments.length;

  // Consumer onboarding incentive (BWP 13 each, cap 50)
  const consumers = onboardingData.filter(o => o.agent_name === merchantName);
  const cappedConsumers = Math.min(consumers.length, 50);
  const onboardingIncentive = cappedConsumers * 13;

  // BWP 95 trigger: >= 2 included transactions
  const bwp95Eligible = txnCount >= 2;

  // Transaction processing incentive (capped at 150 txns)
  const cappedTxns = Math.min(txnCount, 150);
  const fixedIncentive = cappedTxns * 1;
  const variableIncentive = sentPayments.slice(0, 150).reduce((sum, p) => {
    return sum + Math.min(p.amount * 0.005, 27);
  }, 0);

  // Milestone incentive
  let milestoneIncentive = 0;
  let milestonesHit = [];
  if (txnCount >= 50) { milestoneIncentive += 100; milestonesHit.push(50); }
  if (txnCount >= 100) { milestoneIncentive += 125; milestonesHit.push(100); }
  if (txnCount >= 150) { milestoneIncentive += 150; milestonesHit.push(150); }

  const totalIncentive = onboardingIncentive + (bwp95Eligible ? 95 : 0) + fixedIncentive + variableIncentive + milestoneIncentive;

  return {
    merchantName,
    consumers,
    consumerCount: consumers.length,
    cappedConsumers,
    onboardingIncentive,
    bwp95Eligible,
    txnCount,
    cappedTxns,
    fixedIncentive,
    variableIncentive: Math.round(variableIncentive * 100) / 100,
    milestoneIncentive,
    milestonesHit,
    totalIncentive: Math.round(totalIncentive * 100) / 100,
    includedPayments,
    sentPayments,
  };
}

// ─── AUTH ────────────────────────────────────────────────────────────────────
const ADMIN_PASSWORD = "exto@pilot2026";

function LoginScreen({ onLogin }) {
  const [pw, setPw] = useState("");
  const [error, setError] = useState(false);
  const [shake, setShake] = useState(false);

  const handleSubmit = () => {
    if (pw === ADMIN_PASSWORD) {
      onLogin();
    } else {
      setError(true);
      setShake(true);
      setTimeout(() => setShake(false), 500);
    }
  };

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #0f172a 0%, #1e3a5f 60%, #0f4c81 100%)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif" }}>
      <div style={{ background: "rgba(255,255,255,0.05)", backdropFilter: "blur(20px)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 20, padding: "48px 40px", width: 380, animation: shake ? "shake 0.4s ease" : "none" }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ width: 56, height: 56, background: "linear-gradient(135deg, #3b82f6, #06b6d4)", borderRadius: 14, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px", fontSize: 24 }}>🌍</div>
          <div style={{ color: "#fff", fontSize: 22, fontWeight: 700, letterSpacing: -0.5 }}>Exto Pilot</div>
          <div style={{ color: "#94a3b8", fontSize: 13, marginTop: 4 }}>Botswana · Payment Dashboard</div>
        </div>
        <div style={{ marginBottom: 16 }}>
          <label style={{ color: "#94a3b8", fontSize: 12, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1, display: "block", marginBottom: 8 }}>Access Password</label>
          <input
            type="password"
            value={pw}
            onChange={e => { setPw(e.target.value); setError(false); }}
            onKeyDown={e => e.key === "Enter" && handleSubmit()}
            placeholder="Enter team password"
            style={{ width: "100%", padding: "12px 16px", background: "rgba(255,255,255,0.08)", border: error ? "1px solid #ef4444" : "1px solid rgba(255,255,255,0.15)", borderRadius: 10, color: "#fff", fontSize: 15, outline: "none", boxSizing: "border-box" }}
          />
          {error && <div style={{ color: "#ef4444", fontSize: 12, marginTop: 6 }}>Incorrect password. Please try again.</div>}
        </div>
        <button onClick={handleSubmit} style={{ width: "100%", padding: "13px", background: "linear-gradient(135deg, #3b82f6, #06b6d4)", border: "none", borderRadius: 10, color: "#fff", fontSize: 15, fontWeight: 600, cursor: "pointer", letterSpacing: 0.3 }}>
          Access Dashboard →
        </button>
        <div style={{ textAlign: "center", color: "#475569", fontSize: 12, marginTop: 20 }}>
          Pilot Period: 4 months · Botswana Region
        </div>
      </div>
      <style>{`@keyframes shake { 0%,100%{transform:translateX(0)} 25%{transform:translateX(-8px)} 75%{transform:translateX(8px)} }`}</style>
    </div>
  );
}

// ─── COLOUR TOKENS ───────────────────────────────────────────────────────────
const C = {
  bg: "#0f172a", surface: "#1e293b", card: "#1e2d3d", border: "#334155",
  accent: "#3b82f6", cyan: "#06b6d4", green: "#10b981", amber: "#f59e0b", red: "#ef4444",
  text: "#f1f5f9", muted: "#94a3b8", label: "#64748b",
};

// ─── SMALL COMPONENTS ────────────────────────────────────────────────────────
const Badge = ({ color, children }) => (
  <span style={{ background: color + "22", color, border: `1px solid ${color}44`, borderRadius: 6, padding: "2px 8px", fontSize: 11, fontWeight: 600 }}>{children}</span>
);

const KPICard = ({ label, value, sub, color = C.accent, icon }) => (
  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: "20px 24px", flex: 1, minWidth: 160 }}>
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
      <span style={{ fontSize: 20 }}>{icon}</span>
      <span style={{ color: C.muted, fontSize: 12, fontWeight: 500, textTransform: "uppercase", letterSpacing: 0.8 }}>{label}</span>
    </div>
    <div style={{ color, fontSize: 28, fontWeight: 700, letterSpacing: -1 }}>{value}</div>
    {sub && <div style={{ color: C.label, fontSize: 12, marginTop: 4 }}>{sub}</div>}
  </div>
);

const SectionTitle = ({ children }) => (
  <div style={{ color: C.text, fontSize: 16, fontWeight: 700, marginBottom: 16, letterSpacing: -0.3 }}>{children}</div>
);

// ─── ALERT PANEL ─────────────────────────────────────────────────────────────
function AlertPanel({ alerts }) {
  if (!alerts.length) return (
    <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 20, color: C.muted, fontSize: 13 }}>
      ✅ No pending alerts
    </div>
  );
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      {alerts.map((a, i) => (
        <div key={i} style={{ background: a.color + "15", border: `1px solid ${a.color}33`, borderRadius: 10, padding: "12px 16px", display: "flex", gap: 12, alignItems: "flex-start" }}>
          <span style={{ fontSize: 16 }}>{a.icon}</span>
          <div>
            <div style={{ color: a.color, fontSize: 13, fontWeight: 600 }}>{a.merchant}</div>
            <div style={{ color: C.muted, fontSize: 12, marginTop: 2 }}>{a.message}</div>
          </div>
          <Badge color={a.color}>{a.amount}</Badge>
        </div>
      ))}
    </div>
  );
}

// ─── MERCHANT DETAIL ROW ─────────────────────────────────────────────────────
function MerchantRow({ m, expanded, onToggle }) {
  const statusColor = m.bwp95Eligible ? C.green : C.amber;
  const milestone = m.txnCount >= 150 ? "150 ✓" : m.txnCount >= 100 ? "100 ✓" : m.txnCount >= 50 ? "50 ✓" : `${m.txnCount}/50`;

  return (
    <div style={{ border: `1px solid ${expanded ? C.accent + "66" : C.border}`, borderRadius: 12, overflow: "hidden", transition: "border-color 0.2s" }}>
      {/* Row Header */}
      <div onClick={onToggle} style={{ display: "flex", alignItems: "center", padding: "14px 20px", cursor: "pointer", background: expanded ? C.accent + "0a" : C.card, gap: 16 }}>
        <div style={{ width: 36, height: 36, borderRadius: 10, background: `linear-gradient(135deg, ${C.accent}, ${C.cyan})`, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 700, fontSize: 14, flexShrink: 0 }}>
          {m.merchantName.charAt(0).toUpperCase()}
        </div>
        <div style={{ flex: 1 }}>
          <div style={{ color: C.text, fontWeight: 600, fontSize: 14 }}>{m.merchantName}</div>
          <div style={{ color: C.muted, fontSize: 12, marginTop: 2 }}>{m.consumerCount} consumers · {m.txnCount} transactions</div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <Badge color={statusColor}>{m.bwp95Eligible ? "BWP 95 ✓" : "BWP 95 pending"}</Badge>
          <Badge color={C.cyan}>BWP {m.totalIncentive.toLocaleString()}</Badge>
          <span style={{ color: C.muted, fontSize: 18 }}>{expanded ? "▲" : "▼"}</span>
        </div>
      </div>

      {/* Expanded Detail */}
      {expanded && (
        <div style={{ background: "#0f1e2e", padding: "20px 24px", borderTop: `1px solid ${C.border}` }}>
          {/* Incentive Breakdown */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12, marginBottom: 24 }}>
            <div style={{ background: C.card, borderRadius: 10, padding: 16, border: `1px solid ${C.border}` }}>
              <div style={{ color: C.muted, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.8 }}>Consumer Onboarding</div>
              <div style={{ color: C.green, fontSize: 22, fontWeight: 700, margin: "6px 0" }}>BWP {m.onboardingIncentive}</div>
              <div style={{ color: C.label, fontSize: 11 }}>{m.cappedConsumers} consumers × BWP 13{m.consumerCount > 50 ? ` (capped at 50)` : ""}</div>
            </div>
            <div style={{ background: C.card, borderRadius: 10, padding: 16, border: `1px solid ${m.bwp95Eligible ? C.green + "66" : C.amber + "66"}` }}>
              <div style={{ color: C.muted, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.8 }}>Merchant Activation</div>
              <div style={{ color: m.bwp95Eligible ? C.green : C.amber, fontSize: 22, fontWeight: 700, margin: "6px 0" }}>BWP {m.bwp95Eligible ? 95 : 0}</div>
              <div style={{ color: C.label, fontSize: 11 }}>{m.bwp95Eligible ? `Eligible (${m.txnCount} txns completed)` : `Needs ${2 - m.txnCount} more transaction(s)`}</div>
            </div>
            <div style={{ background: C.card, borderRadius: 10, padding: 16, border: `1px solid ${C.border}` }}>
              <div style={{ color: C.muted, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.8 }}>Txn Processing (Fixed)</div>
              <div style={{ color: C.accent, fontSize: 22, fontWeight: 700, margin: "6px 0" }}>BWP {m.fixedIncentive}</div>
              <div style={{ color: C.label, fontSize: 11 }}>{m.cappedTxns} txns × BWP 1{m.txnCount > 150 ? " (capped at 150)" : ""}</div>
            </div>
            <div style={{ background: C.card, borderRadius: 10, padding: 16, border: `1px solid ${C.border}` }}>
              <div style={{ color: C.muted, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.8 }}>Txn Processing (Variable)</div>
              <div style={{ color: C.cyan, fontSize: 22, fontWeight: 700, margin: "6px 0" }}>BWP {m.variableIncentive.toFixed(2)}</div>
              <div style={{ color: C.label, fontSize: 11 }}>0.5% per txn, max BWP 27 each</div>
            </div>
            <div style={{ background: C.card, borderRadius: 10, padding: 16, border: `1px solid ${m.milestonesHit.length ? C.amber + "66" : C.border}` }}>
              <div style={{ color: C.muted, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.8 }}>Milestone Bonus</div>
              <div style={{ color: C.amber, fontSize: 22, fontWeight: 700, margin: "6px 0" }}>BWP {m.milestoneIncentive}</div>
              <div style={{ color: C.label, fontSize: 11 }}>
                {m.milestonesHit.length ? `Hit: ${m.milestonesHit.join(", ")} txns` : `Next: 50 txns → BWP 100`}
              </div>
            </div>
            <div style={{ background: `linear-gradient(135deg, ${C.accent}22, ${C.cyan}22)`, borderRadius: 10, padding: 16, border: `1px solid ${C.accent}44` }}>
              <div style={{ color: C.muted, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.8 }}>Total Payable</div>
              <div style={{ color: "#fff", fontSize: 22, fontWeight: 700, margin: "6px 0" }}>BWP {m.totalIncentive.toLocaleString()}</div>
              <div style={{ color: C.label, fontSize: 11 }}>All incentives combined</div>
            </div>
          </div>

          {/* Milestone progress */}
          <div style={{ marginBottom: 20 }}>
            <div style={{ color: C.muted, fontSize: 12, marginBottom: 8, fontWeight: 600 }}>TRANSACTION MILESTONES</div>
            <div style={{ display: "flex", gap: 12 }}>
              {[{ n: 50, reward: 100 }, { n: 100, reward: 125 }, { n: 150, reward: 150 }].map(ms => {
                const hit = m.txnCount >= ms.n;
                return (
                  <div key={ms.n} style={{ flex: 1, background: hit ? C.amber + "22" : C.card, border: `1px solid ${hit ? C.amber + "66" : C.border}`, borderRadius: 8, padding: "10px 12px", textAlign: "center" }}>
                    <div style={{ color: hit ? C.amber : C.muted, fontWeight: 700, fontSize: 15 }}>{ms.n} txns</div>
                    <div style={{ color: hit ? C.amber : C.label, fontSize: 12, marginTop: 2 }}>BWP {ms.reward} {hit ? "✓" : ""}</div>
                    {!hit && <div style={{ color: C.label, fontSize: 11, marginTop: 2 }}>{ms.n - m.txnCount} more needed</div>}
                  </div>
                );
              })}
            </div>
          </div>

          {/* Consumers list */}
          {m.consumers.length > 0 && (
            <div style={{ marginBottom: 20 }}>
              <div style={{ color: C.muted, fontSize: 12, marginBottom: 8, fontWeight: 600 }}>ONBOARDED CONSUMERS ({m.consumers.length})</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                {m.consumers.map((c, i) => (
                  <span key={i} style={{ background: C.green + "18", border: `1px solid ${C.green}33`, color: C.green, borderRadius: 6, padding: "3px 10px", fontSize: 12 }}>{c.user_name}</span>
                ))}
              </div>
            </div>
          )}

          {/* Recent transactions */}
          {m.includedPayments.length > 0 && (
            <div>
              <div style={{ color: C.muted, fontSize: 12, marginBottom: 8, fontWeight: 600 }}>RECENT TRANSACTIONS (showing up to 10)</div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead>
                    <tr>
                      {["Date", "Sender", "Receiver", "Amount (BWP)", "Incentive (BWP)"].map(h => (
                        <th key={h} style={{ color: C.label, padding: "6px 12px", textAlign: "left", borderBottom: `1px solid ${C.border}`, fontWeight: 600 }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {m.includedPayments.slice(0, 10).map((p, i) => {
                      const incentive = Math.min(p.amount * 0.005, 27);
                      return (
                        <tr key={i} style={{ borderBottom: `1px solid ${C.border}22` }}>
                          <td style={{ padding: "8px 12px", color: C.muted }}>{p.created_at.split(",")[0]}</td>
                          <td style={{ padding: "8px 12px", color: C.text }}>{p.sender || "—"}</td>
                          <td style={{ padding: "8px 12px", color: C.text }}>{p.receiver || "—"}</td>
                          <td style={{ padding: "8px 12px", color: C.cyan, fontWeight: 600 }}>{p.amount.toLocaleString()}</td>
                          <td style={{ padding: "8px 12px", color: C.green }}>{incentive.toFixed(2)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── MAIN DASHBOARD ──────────────────────────────────────────────────────────
function Dashboard({ onboardingData, paymentsData, dataSource, lastUpdated, onRefresh, refreshing }) {
  const [activeTab, setActiveTab] = useState("overview");
  const [expandedMerchant, setExpandedMerchant] = useState(null);
  const [searchTerm, setSearchTerm] = useState("");

  // Compute all merchant profiles
  const merchantNames = useMemo(() => {
    const names = new Set();
    onboardingData.forEach(o => { if (o.agent_name) names.add(o.agent_name); });
    paymentsData.forEach(p => { if (p.sender) names.add(p.sender); });
    return [...names].sort();
  }, []);

  const merchantProfiles = useMemo(() =>
    merchantNames.map(name => calcMerchantIncentives(name, onboardingData, paymentsData)),
    [merchantNames]
  );

  // KPI totals
  const kpis = useMemo(() => {
    const totalConsumers = new Set(onboardingData.filter(o => o.agent_name).map(o => o.user_name)).size;
    const totalMerchants = merchantNames.length;
    const includedPayments = paymentsData.filter(p => p.included);
    const totalVolume = includedPayments.reduce((s, p) => s + p.amount, 0);
    const totalTxns = includedPayments.length;
    const totalIncentive = merchantProfiles.reduce((s, m) => s + m.totalIncentive, 0);
    return { totalConsumers, totalMerchants, totalVolume, totalTxns, totalIncentive };
  }, [merchantProfiles]);

  // Alerts
  const alerts = useMemo(() => {
    const list = [];
    merchantProfiles.forEach(m => {
      if (m.bwp95Eligible) list.push({ merchant: m.merchantName, message: "Completed 2+ transactions — BWP 95 activation bonus payable", amount: "BWP 95", color: C.green, icon: "✅" });
      m.milestonesHit.forEach(ms => list.push({ merchant: m.merchantName, message: `${ms}-transaction milestone reached`, amount: `BWP ${ms === 50 ? 100 : ms === 100 ? 125 : 150}`, color: C.amber, icon: "🏆" }));
      if (m.consumerCount >= 50) list.push({ merchant: m.merchantName, message: "Consumer onboarding cap of 50 reached", amount: "Cap hit", color: C.red, icon: "⚠️" });
    });
    return list;
  }, [merchantProfiles]);

  // Chart data
  const txnVolumeByDate = useMemo(() => {
    const byDate = {};
    paymentsData.filter(p => p.included).forEach(p => {
      const date = p.created_at.split(",")[0];
      byDate[date] = (byDate[date] || 0) + p.amount;
    });
    return Object.entries(byDate).sort((a, b) => new Date(a[0]) - new Date(b[0])).map(([date, vol]) => ({ date, volume: Math.round(vol) }));
  }, []);

  const onboardingByDate = useMemo(() => {
    const byDate = {};
    onboardingData.forEach(o => {
      const date = o.created_at.split(",")[0];
      byDate[date] = (byDate[date] || 0) + 1;
    });
    return Object.entries(byDate).sort().map(([date, count]) => ({ date, count }));
  }, []);

  const leaderboard = useMemo(() =>
    [...merchantProfiles].sort((a, b) => b.txnCount - a.txnCount).slice(0, 8),
    [merchantProfiles]
  );

  const incentiveBreakdown = useMemo(() => {
    const totals = merchantProfiles.reduce((acc, m) => ({
      onboarding: acc.onboarding + m.onboardingIncentive + (m.bwp95Eligible ? 95 : 0),
      fixed: acc.fixed + m.fixedIncentive,
      variable: acc.variable + m.variableIncentive,
      milestone: acc.milestone + m.milestoneIncentive,
    }), { onboarding: 0, fixed: 0, variable: 0, milestone: 0 });
    return [
      { name: "Onboarding", value: Math.round(totals.onboarding), color: C.green },
      { name: "Fixed Txn", value: Math.round(totals.fixed), color: C.accent },
      { name: "Variable Txn", value: Math.round(totals.variable), color: C.cyan },
      { name: "Milestones", value: Math.round(totals.milestone), color: C.amber },
    ];
  }, [merchantProfiles]);

  const filteredMerchants = merchantProfiles.filter(m =>
    m.merchantName.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const tabs = [
    { id: "overview", label: "📊 Overview" },
    { id: "merchants", label: "🏪 Merchants" },
    { id: "alerts", label: `🔔 Alerts ${alerts.length ? `(${alerts.length})` : ""}` },
  ];

  return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'Inter', system-ui, sans-serif", color: C.text }}>
      {/* Top Bar */}
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: "0 32px", display: "flex", alignItems: "center", height: 60 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flex: 1 }}>
          <span style={{ fontSize: 22 }}>🌍</span>
          <div>
            <span style={{ fontWeight: 700, fontSize: 16 }}>Exto Pilot</span>
            <span style={{ color: C.muted, fontSize: 13, marginLeft: 8 }}>Botswana · Admin Dashboard</span>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 14, marginRight: 16 }}>
          <Badge color={dataSource === "live" ? C.green : C.amber}>
            {dataSource === "live" ? "🟢 Live OneDrive Data" : "🟡 Sample Data"}
          </Badge>
          {lastUpdated && (
            <span style={{ color: C.label, fontSize: 12 }}>Updated {lastUpdated}</span>
          )}
          <button
            onClick={onRefresh}
            disabled={refreshing}
            style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 8, padding: "6px 12px", color: C.text, fontSize: 12, fontWeight: 500, cursor: refreshing ? "default" : "pointer", display: "flex", alignItems: "center", gap: 6, opacity: refreshing ? 0.6 : 1 }}
          >
            <span style={{ display: "inline-block", animation: refreshing ? "spin 1s linear infinite" : "none" }}>🔄</span>
            {refreshing ? "Refreshing..." : "Refresh"}
          </button>
        </div>
        <div style={{ display: "flex", gap: 6 }}>
          {tabs.map(t => (
            <button key={t.id} onClick={() => setActiveTab(t.id)} style={{ background: activeTab === t.id ? C.accent : "transparent", color: activeTab === t.id ? "#fff" : C.muted, border: "none", borderRadius: 8, padding: "6px 14px", fontSize: 13, fontWeight: 500, cursor: "pointer" }}>
              {t.label}
            </button>
          ))}
        </div>
      </div>
      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>

      <div style={{ padding: "28px 32px", maxWidth: 1400, margin: "0 auto" }}>

        {/* ── OVERVIEW TAB ── */}
        {activeTab === "overview" && (
          <div>
            {/* KPI Cards */}
            <div style={{ display: "flex", gap: 16, marginBottom: 28, flexWrap: "wrap" }}>
              <KPICard icon="🏪" label="Merchants" value={kpis.totalMerchants} sub="Active in pilot" color={C.accent} />
              <KPICard icon="👥" label="Consumers Onboarded" value={kpis.totalConsumers} sub="Unique verified consumers" color={C.green} />
              <KPICard icon="⚡" label="Transactions" value={kpis.totalTxns.toLocaleString()} sub="Included payments" color={C.cyan} />
              <KPICard icon="💰" label="Total Volume" value={`BWP ${Math.round(kpis.totalVolume).toLocaleString()}`} sub="Across all payments" color={C.amber} />
              <KPICard icon="🎯" label="Total Incentive Liability" value={`BWP ${Math.round(kpis.totalIncentive).toLocaleString()}`} sub="All merchants combined" color={C.red} />
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginBottom: 20 }}>
              {/* Transaction Volume Chart */}
              <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 24 }}>
                <SectionTitle>Transaction Volume Over Time (BWP)</SectionTitle>
                <ResponsiveContainer width="100%" height={200}>
                  <BarChart data={txnVolumeByDate}>
                    <XAxis dataKey="date" tick={{ fill: C.label, fontSize: 10 }} />
                    <YAxis tick={{ fill: C.label, fontSize: 10 }} />
                    <Tooltip contentStyle={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, color: C.text }} />
                    <Bar dataKey="volume" fill={C.accent} radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>

              {/* Onboarding Trend */}
              <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 24 }}>
                <SectionTitle>Onboarding Trend</SectionTitle>
                <ResponsiveContainer width="100%" height={200}>
                  <LineChart data={onboardingByDate}>
                    <XAxis dataKey="date" tick={{ fill: C.label, fontSize: 10 }} />
                    <YAxis tick={{ fill: C.label, fontSize: 10 }} />
                    <Tooltip contentStyle={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, color: C.text }} />
                    <Line type="monotone" dataKey="count" stroke={C.green} strokeWidth={2} dot={{ fill: C.green, r: 3 }} />
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr", gap: 20, marginBottom: 20 }}>
              {/* Merchant Leaderboard */}
              <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 24 }}>
                <SectionTitle>Merchant Leaderboard — by Transactions</SectionTitle>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {leaderboard.map((m, i) => (
                    <div key={m.merchantName} style={{ display: "flex", alignItems: "center", gap: 12 }}>
                      <span style={{ color: i < 3 ? C.amber : C.label, fontWeight: 700, fontSize: 13, width: 20 }}>#{i + 1}</span>
                      <span style={{ flex: 1, color: C.text, fontSize: 13 }}>{m.merchantName}</span>
                      <div style={{ width: 120, background: C.border, borderRadius: 4, height: 6, overflow: "hidden" }}>
                        <div style={{ width: `${Math.min((m.txnCount / (leaderboard[0].txnCount || 1)) * 100, 100)}%`, background: i === 0 ? C.amber : C.accent, height: "100%", borderRadius: 4 }} />
                      </div>
                      <span style={{ color: C.cyan, fontSize: 13, fontWeight: 600, width: 50, textAlign: "right" }}>{m.txnCount}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Incentive Breakdown Pie */}
              <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 24 }}>
                <SectionTitle>Incentive Breakdown</SectionTitle>
                <ResponsiveContainer width="100%" height={180}>
                  <PieChart>
                    <Pie data={incentiveBreakdown} cx="50%" cy="50%" innerRadius={50} outerRadius={75} paddingAngle={3} dataKey="value">
                      {incentiveBreakdown.map((entry, i) => <Cell key={i} fill={entry.color} />)}
                    </Pie>
                    <Tooltip contentStyle={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, color: C.text }} formatter={(v) => `BWP ${v}`} />
                    <Legend formatter={(v) => <span style={{ color: C.muted, fontSize: 11 }}>{v}</span>} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Pilot Progress */}
            <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 24 }}>
              <SectionTitle>Pilot Progress Trackers</SectionTitle>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 20 }}>
                {[
                  { label: "Consumers Onboarded (Pilot Target)", current: kpis.totalConsumers, target: 200, color: C.green },
                  { label: "Total Transactions (Platform Target)", current: kpis.totalTxns, target: 500, color: C.accent },
                  { label: "Incentive Deployed (BWP Budget)", current: Math.round(kpis.totalIncentive), target: 50000, color: C.amber },
                ].map(p => {
                  const pct = Math.min((p.current / p.target) * 100, 100);
                  return (
                    <div key={p.label}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                        <span style={{ color: C.muted, fontSize: 12 }}>{p.label}</span>
                        <span style={{ color: p.color, fontSize: 12, fontWeight: 600 }}>{Math.round(pct)}%</span>
                      </div>
                      <div style={{ background: C.border, borderRadius: 8, height: 8 }}>
                        <div style={{ width: `${pct}%`, background: p.color, borderRadius: 8, height: "100%", transition: "width 1s ease" }} />
                      </div>
                      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4 }}>
                        <span style={{ color: C.text, fontSize: 12, fontWeight: 600 }}>{p.current.toLocaleString()}</span>
                        <span style={{ color: C.label, fontSize: 12 }}>Target: {p.target.toLocaleString()}</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {/* ── MERCHANTS TAB ── */}
        {activeTab === "merchants" && (
          <div>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20 }}>
              <div style={{ color: C.text, fontSize: 20, fontWeight: 700 }}>All Merchants <span style={{ color: C.muted, fontSize: 14, fontWeight: 400 }}>({filteredMerchants.length})</span></div>
              <input
                placeholder="🔍  Search merchant..."
                value={searchTerm}
                onChange={e => setSearchTerm(e.target.value)}
                style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 10, padding: "8px 16px", color: C.text, fontSize: 13, outline: "none", width: 240 }}
              />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {filteredMerchants.map(m => (
                <MerchantRow
                  key={m.merchantName}
                  m={m}
                  expanded={expandedMerchant === m.merchantName}
                  onToggle={() => setExpandedMerchant(expandedMerchant === m.merchantName ? null : m.merchantName)}
                />
              ))}
              {filteredMerchants.length === 0 && (
                <div style={{ textAlign: "center", color: C.muted, padding: 40 }}>No merchants match your search.</div>
              )}
            </div>
          </div>
        )}

        {/* ── ALERTS TAB ── */}
        {activeTab === "alerts" && (
          <div>
            <div style={{ color: C.text, fontSize: 20, fontWeight: 700, marginBottom: 20 }}>
              Pending Alerts <span style={{ color: C.muted, fontSize: 14, fontWeight: 400 }}>({alerts.length} items)</span>
            </div>
            <AlertPanel alerts={alerts} />
            {alerts.length > 0 && (
              <div style={{ marginTop: 16, background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: 16 }}>
                <div style={{ color: C.muted, fontSize: 12, marginBottom: 8, fontWeight: 600 }}>ALERT SUMMARY</div>
                <div style={{ display: "flex", gap: 20 }}>
                  <div><span style={{ color: C.green, fontWeight: 700 }}>{alerts.filter(a => a.color === C.green).length}</span> <span style={{ color: C.muted, fontSize: 13 }}>BWP 95 payable</span></div>
                  <div><span style={{ color: C.amber, fontWeight: 700 }}>{alerts.filter(a => a.color === C.amber).length}</span> <span style={{ color: C.muted, fontSize: 13 }}>Milestone bonuses</span></div>
                  <div><span style={{ color: C.red, fontWeight: 700 }}>{alerts.filter(a => a.color === C.red).length}</span> <span style={{ color: C.muted, fontSize: 13 }}>Cap warnings</span></div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── MICROSOFT SIGN-IN SCREEN (after password gate) ─────────────────────────
function MicrosoftSignIn({ onUseSampleData }) {
  const { instance } = useMsal();
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleSignIn = async () => {
    setLoading(true);
    setError(null);
    try {
      await instance.loginPopup(loginRequest);
    } catch (err) {
      setError(err.message || "Sign-in failed. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #0f172a 0%, #1e3a5f 60%, #0f4c81 100%)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif" }}>
      <div style={{ background: "rgba(255,255,255,0.05)", backdropFilter: "blur(20px)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 20, padding: "48px 40px", width: 420, textAlign: "center" }}>
        <div style={{ width: 56, height: 56, background: "linear-gradient(135deg, #3b82f6, #06b6d4)", borderRadius: 14, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px", fontSize: 24 }}>🌍</div>
        <div style={{ color: "#fff", fontSize: 20, fontWeight: 700, marginBottom: 8 }}>Connect to Live Data</div>
        <div style={{ color: "#94a3b8", fontSize: 13, marginBottom: 28, lineHeight: 1.5 }}>
          Sign in with your Microsoft work account to pull live data from the pilot Excel file on OneDrive.
        </div>
        <button onClick={handleSignIn} disabled={loading} style={{ width: "100%", padding: "13px", background: "#fff", border: "none", borderRadius: 10, color: "#1e293b", fontSize: 14, fontWeight: 600, cursor: "pointer", marginBottom: 12, opacity: loading ? 0.7 : 1 }}>
          {loading ? "Signing in..." : "🔑  Sign in with Microsoft"}
        </button>
        <button onClick={onUseSampleData} style={{ width: "100%", padding: "11px", background: "transparent", border: "1px solid rgba(255,255,255,0.15)", borderRadius: 10, color: "#94a3b8", fontSize: 13, cursor: "pointer" }}>
          Continue with sample data instead
        </button>
        {error && <div style={{ color: "#ef4444", fontSize: 12, marginTop: 14 }}>{error}</div>}
      </div>
    </div>
  );
}

// ─── LIVE DATA ORCHESTRATOR ──────────────────────────────────────────────────
function LiveDataApp() {
  const isAuthenticated = useIsAuthenticated();
  const [useSample, setUseSample] = useState(false);
  const [onboarding, setOnboarding] = useState(FALLBACK_ONBOARDING);
  const [payments, setPayments] = useState(FALLBACK_PAYMENTS);
  const [dataSource, setDataSource] = useState("sample");
  const [lastUpdated, setLastUpdated] = useState(null);
  const [refreshing, setRefreshing] = useState(false);
  const [fetchError, setFetchError] = useState(null);
  const intervalRef = useRef(null);

  const loadLiveData = useCallback(async () => {
    setRefreshing(true);
    setFetchError(null);
    try {
      const result = await fetchLivePilotData();
      setOnboarding(result.onboarding);
      setPayments(result.payments);
      setDataSource("live");
      setLastUpdated(new Date(result.fetchedAt).toLocaleTimeString());
    } catch (err) {
      console.error("Live data fetch failed:", err);
      setFetchError(err.message);
      // Keep showing whatever data we had before (graceful degradation)
    } finally {
      setRefreshing(false);
    }
  }, []);

  useEffect(() => {
    if (!isAuthenticated) return;
    loadLiveData();
    intervalRef.current = setInterval(loadLiveData, AUTO_REFRESH_MS);
    return () => clearInterval(intervalRef.current);
  }, [isAuthenticated, loadLiveData]);

  if (!isAuthenticated && !useSample) {
    return <MicrosoftSignIn onUseSampleData={() => setUseSample(true)} />;
  }

  return (
    <Dashboard
      onboardingData={onboarding}
      paymentsData={payments}
      dataSource={dataSource}
      lastUpdated={lastUpdated}
      onRefresh={loadLiveData}
      refreshing={refreshing}
    />
  );
}

// ─── APP ROOT ────────────────────────────────────────────────────────────────
export default function App() {
  const [loggedIn, setLoggedIn] = useState(false);
  if (!loggedIn) return <LoginScreen onLogin={() => setLoggedIn(true)} />;
  return (
    <MsalProvider instance={msalInstance}>
      <LiveDataApp />
    </MsalProvider>
  );
}
